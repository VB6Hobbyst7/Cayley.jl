Attribute VB_Name = "modCreditUsageSheet"
Option Explicit
Public g_StartRunCreditUsageSheet As Double

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuCreditUsageSheet
' Author    : Philip Swannell
' Date      : 22-May-2015
' Purpose   : Attached to "Menu..." button
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuCreditUsageSheet(Optional ByVal Choice As String)
          
1         On Error GoTo ErrHandler
          
          Dim EnableFlags
          Dim FaceIDs
          Dim LinesBookIsOpen As Boolean
          Dim MarketBookIsOpen As Boolean
          Dim TheChoices
          Dim TradesBookIsOpen As Boolean

          Const chFeedRates = "Feed &Rates to Market Data Workbook"
          Const FidFeedRates = 4355
          Const chRebuildModel = "&Build Hull-White Model"
          Const FidRebuildModel = 5828
          Const chCalc = "Calculate P&FE                       (Shift F9)"
          Const FidCalc = 283
          Dim chSolveTHR As String
2         chSolveTHR = "Solve for Trade Headroom   (1 to " & GetHedgeHorizon() & "Y, &Either or basis)"
          Const FidSolveTHR = 156
          Const chSolveTHR2 = "Solve for Trade Headroom   (&Joint 3, 4 and 5Y)"
          Const FidSolveTHR2 = 156
          Const chSolveTHR3 = "Solve for Trade Headroom   (&Custom allocation)..."
          Const FidSolveTHR3 = 156
          Const chCreateSystemImage = "&Create Julia System Image..."
          Const FidCreateSystemImage = 11208
          Const chLaunchJuliaWithoutSystemImage = "&Launch Julia without System Image"
          Const FidLaunchJuliaWithoutSystemImage = 16330
          Const chSolveFxHR = "Solve for F&x Headroom"
          Const FidSolveFxHR = 384
          Const chSolveFxVolHR = "Solve for Fx &Vol Headroom"
          Const FidSolveFxVolHR = 7707
          Const chShowExtraTrades = "Show current ""E&Xtra Trades"" in Julia format"
          Const FidShowExtraTrades = 0
          
          Const chPaste = "&Paste Charts to new workbook..."
          Const FidPaste = 422
          Dim Allocations As Variant
          Dim chOpenLines As String
          Dim chOpenMarket As String
          Dim chOpenOthers As Variant
          Dim chOpenTrades As String
          Dim EnableOpenOthers As Variant
          Dim FidOpenLines As Long
          Dim FidOpenMarket As Long
          Dim FidOpenOthers As Variant
          Dim FidOpenTrades As Long
          Const chVersionInfo As String = "Show &Version Info..."
          Const FidVersionInfo = 487
          
          Const chShowTrades = "&Show trades and PVs"
          Const FidShowTrades = 1987
          Dim AdvancedChoices
          Dim OBAO As Boolean
          Dim SPH As SolumAddin.clsScreenUpdateHandler

          Dim Activate As Boolean
          Dim TWIOOD As Boolean

          Dim TradesWorkbookName As String
          
3         RunThisAtTopOfCallStack

4         OBAO = OtherBooksAreOpen(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen)
                   
5         If TradesBookIsOpen Then
6             TWIOOD = TradesWorkbookIsOutOfDate()
7         End If

8         TradesWorkbookName = gCayleyTradesWorkbookName

9         If Choice = "" Then        'Build the menu
10            chOpenOthers = NameForOpenOthers(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen, True)
11            FidOpenOthers = IIf(IsMissing(chOpenOthers), createmissing(), 23)
12            EnableOpenOthers = IIf(IsMissing(chOpenOthers), createmissing(), True)
              
13            If TradesBookIsOpen Then
14                If TWIOOD Then
15                    chOpenTrades = "Re-open &Trade files (they have changed!)"
16                    FidOpenTrades = 16368
17                Else
18                    chOpenTrades = "Activate &Trades files"
19                    FidOpenTrades = 142
20                End If
21            Else
22                chOpenTrades = "Open &Trades files              (Shift to View)"
23                FidOpenTrades = 23
24            End If

25            If LinesBookIsOpen Then
26                chOpenLines = "Activate &Lines workbook"
27                FidOpenLines = 142
28            Else
29                chOpenLines = "Open &Lines workbook                (Shift to Activate)"
30                FidOpenLines = 23
31            End If
32            If MarketBookIsOpen Then
33                MarketBookIsOpen = True
34                chOpenMarket = "Activate &Market Data workbook"
35                FidOpenMarket = 142
36            Else
37                chOpenMarket = "Open &Market Data workbook    (Shift to Activate)"
38                FidOpenMarket = 23
39            End If

40            TheChoices = sArrayStack(chFeedRates, chRebuildModel, "--" & chCalc, chSolveTHR, chSolveTHR2, _
                  chSolveTHR3, chSolveFxHR, chSolveFxVolHR, "--" & chShowTrades, _
                   chPaste, "--" & chOpenTrades, _
                  chOpenLines, chOpenMarket, chOpenOthers)
41            AdvancedChoices = sArrayRange("--&Developer Tools", _
                  sArrayStack(chVersionInfo, "--" & chCreateSystemImage, chLaunchJuliaWithoutSystemImage, "--" & chShowExtraTrades))
42            TheChoices = sArrayStack(TheChoices, AdvancedChoices)

43            FaceIDs = sArrayStack(FidFeedRates, FidRebuildModel, FidCalc, FidSolveTHR, FidSolveTHR2, _
                  FidSolveTHR3, FidSolveFxHR, FidSolveFxVolHR, FidShowTrades, _
                   FidPaste, FidOpenTrades, _
                  FidOpenLines, FidOpenMarket, FidOpenOthers, FidVersionInfo, FidCreateSystemImage, _
                  FidLaunchJuliaWithoutSystemImage, FidShowExtraTrades)
                  
44            EnableFlags = sArrayStack(OBAO, OBAO, OBAO, OBAO, OBAO, OBAO, OBAO, OBAO, _
                  OBAO, OBAO, True, True, True, EnableOpenOthers, sReshape(True, sNRows(AdvancedChoices), 1))

45            Choice = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , ChooseAnchorObject())
46        End If
47        If Choice = "#Cancel!" Then Exit Sub
48        g_StartRunCreditUsageSheet = sElapsedTime()

49        Set SPH = CreateScreenUpdateHandler()

          'If ever we want to call this method from other methods then add hard-wired _
           string (e.g. "Calculate") to the case statements below
50        Select Case Choice
              Case Unembellish(chFeedRates)
51                FeedRatesFromTextFile
52            Case Unembellish(CStr(chOpenOthers))
53                AddFilters
54                OpenOtherBooks
55            Case Unembellish(chCalc), "Calculate"
56                JuliaLaunchForCayley
57                AddFilters
58                RunCreditUsageSheet "Standard", True, False, True
59            Case Unembellish(chSolveTHR)
60                JuliaLaunchForCayley
61                AddFilters
62                RunCreditUsageSheet "Solve1to5", True, False, True
63            Case Unembellish(chSolveTHR2)
64                JuliaLaunchForCayley
65                AddFilters
66                Allocations = sReshape(0, GetHedgeHorizon(), 1)
67                Allocations(3, 1) = 1 / 3: Allocations(4, 1) = 1 / 3: Allocations(5, 1) = 1 / 3
68                RunCreditUsageSheet "Solve345", True, False, True, , Allocations
69            Case Unembellish(chSolveTHR3)
70                JuliaLaunchForCayley
71                AddFilters
                  Dim Allocation
72                Allocation = GetAllocation()
73                If Not IsEmpty(Allocation) Then
74                    RunCreditUsageSheet "Solve345", True, False, True, , Allocation
75                End If
76            Case Unembellish(chRebuildModel)
77                JuliaLaunchForCayley
78                BuildModelsInJulia True, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")
79            Case Unembellish(chSolveFxVolHR)
80                JuliaLaunchForCayley
81                RunCreditUsageSheet "SolveFxVol", True, False, True
82            Case Unembellish(chSolveFxHR)
83                JuliaLaunchForCayley
84                AddFilters
85                RunCreditUsageSheet "SolveFx", True, False, True
86            Case Unembellish(chOpenLines)
87                Activate = IIf(LinesBookIsOpen, True, IsShiftKeyDown())
88                OpenLinesWorkbook Not (Activate), Activate
89            Case Unembellish(chOpenTrades)
90                Activate = IIf(TradesBookIsOpen, True, IsShiftKeyDown())
91                If TWIOOD Then
92                    LoadTradesFromTextFiles , , , True
93                Else
94                    OpenTradesWorkbook Not (Activate), Activate
95                End If
96            Case Unembellish(chOpenMarket)
97                Activate = IIf(MarketBookIsOpen, True, IsShiftKeyDown())
98                OpenMarketWorkbook Not (Activate), Activate
99            Case Unembellish(chPaste)
100               JuliaLaunchForCayley
101               PasteCharts
102           Case Unembellish(chShowTrades), "ShowTrades"
103               JuliaLaunchForCayley
104               ShowTrades
105           Case Unembellish(chCreateSystemImage)
106               JuliaCreateSystemImage True, UseLinux()
107           Case Unembellish(chLaunchJuliaWithoutSystemImage)
108               LaunchJuliaWithoutSystemImage
109           Case Unembellish(chShowExtraTrades)
110               JuliaLaunchForCayley
111               ShowExtraTrades
112           Case Unembellish(chVersionInfo)
113               ShowVersionInfo
114           Case Else
115               Throw "Unrecognised choice: " & Choice
116       End Select

          'Final tidy up - and we really want protection on the CreditUsage sheet to be on. _
           It's too easy for the user to make inadvertant changes.
117       If Not ActiveSheet Is Nothing Then
118           If ActiveSheet Is shCreditUsage Then
119               UnselectChart
120               FormatCreditUsageSheet True
121           End If
122       End If

          Dim TimingMessage As String
123       If Choice = chCalc Then
124           Choice = "Calculate PFE"
125       Else
126           Choice = Replace(Choice, "...", "")
127       End If
128       TimingMessage = "Time to " & Choice & ": " & _
              Format(sElapsedTime() - g_StartRunCreditUsageSheet, "0.00") & " seconds"
129       TemporaryMessage TimingMessage, , False

130       Exit Sub
ErrHandler:
131       SomethingWentWrong "#MenuCreditUsageSheet (line " & CStr(Erl) & "): " & Err.Description & "!", , "Cayley"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunCreditUsageSheet
' Author    : Philip Swannell
' Date      : 30-Sep-2016
' Purpose   : Grabs data from PFE and Config sheets and calls lower-level RunCreditUsageSheetCore
' DisplayBaseCasePFE means "leave the PFE sheet showing results excluding Extra Trades" or without (additional) Fx shock
' RefreshSheet determines if the sheet is re-drawn.
' -----------------------------------------------------------------------------------------------------------------------
Function RunCreditUsageSheet(Mode As String, ThrowErrors As Boolean, DisplayBaseCasePFE As Boolean, _
          RefreshSheet As Boolean, Optional ByRef ResultsDict As Dictionary, Optional Allocations)
          
          Dim CurrenciesToInclude As String
          Dim ExtraTradeAmounts As Variant
          Dim ExtraTradeLabels As Variant
          Dim ExtraTradesAre As String
          Dim Filter1Value As Variant
          Dim Filter2Value As Variant
          Dim FilterBy1 As String
          Dim FilterBy2 As String
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim IncludeAssetClasses As String
          Dim IncludeExtraTrades As Boolean
          Dim IncludeFutureTrades As Boolean
          Dim LinesScaleFactor As Double
          Dim NumMCPaths As Long
          Dim NumObservations As Double
          Dim PortfolioAgeing As Double
          Dim TradesScaleFactor As Double

1         On Error GoTo ErrHandler
2         If ResultsDict Is Nothing Then Set ResultsDict = New Dictionary
3         Filter1Value = RangeFromSheet(shCreditUsage, "Filter1Value", True, True, True, False, False)
4         FilterBy1 = RangeFromSheet(shCreditUsage, "FilterBy1", False, True, False, False, False)
5         FilterBy2 = RangeFromSheet(shCreditUsage, "FilterBy2", False, True, False, False, False)
6         Filter2Value = RangeFromSheet(shCreditUsage, "Filter2Value", True, True, True, False, False)
7         IncludeFutureTrades = RangeFromSheet(shCreditUsage, "IncludeFutureTrades", False, False, True, False, False)
8         IncludeAssetClasses = RangeFromSheet(shCreditUsage, "IncludeAssetClasses", False, True, False, False, False)
9         IncludeExtraTrades = RangeFromSheet(shCreditUsage, "IncludeExtraTrades", False, False, True, False, False)
10        ExtraTradeLabels = RangeFromSheet(shCreditUsage, "ExtraTradeLabels", False, True, False, False, False)
11        ExtraTradeAmounts = RangeFromSheet(shCreditUsage, "ExtraTradeAmounts", True, False, False, False, False)
12        PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False)
13        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False)
14        CurrenciesToInclude = RangeFromSheet(shConfig, "CurrenciesToInclude", False, True, False, False, False)
15        NumMCPaths = RangeFromSheet(shCreditUsage, "NumMCPaths", True, False, False, False, False)
16        NumObservations = RangeFromSheet(shCreditUsage, "NumObservations", True, False, False, False, False)
17        FxShock = RangeFromSheet(shCreditUsage, "FxShock", True, False, False, False, False)
18        FxVolShock = RangeFromSheet(shCreditUsage, "FxVolShock", True, False, False, False, False)
19        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False)
20        ExtraTradesAre = RangeFromSheet(shCreditUsage, "ExtraTradesAre", False, True, False, False, False)
21        RunCreditUsageSheet = RunCreditUsageSheetCore(Mode, FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, IncludeAssetClasses, IncludeExtraTrades, ExtraTradeLabels, _
              ExtraTradeAmounts, PortfolioAgeing, TradesScaleFactor, CurrenciesToInclude, NumMCPaths, _
              NumObservations, FxShock, FxVolShock, LinesScaleFactor, _
              ThrowErrors, DisplayBaseCasePFE, RefreshSheet, ResultsDict, Allocations, ExtraTradesAre)
22        Exit Function
ErrHandler:
23        Throw "#RunCreditUsageSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunCreditUsageSheetCore
' Author    : Philip Swannell
' Date      : 06-Sep-2016
' Purpose   : According to Mode:
'  Mode = "Standard": Calculates PFE \ Notional-based profile by calling PFEProfileFromFilters
'  Mode = "Solve345": Calculated the trade Headroom for simultaneous 3,4, and 5 year ATM EURUSD forwards
'  Mode = "Solve1to5": Calculates the trade Headroom 1 to 5 years EURUSD forwards either or basis
'  Mode = "SolveFx": Solves for the Fx Headroom
' The method is not yet used for Fx Vol solving - will need to be if we port such solving to R. For
' the time being the method SolveFxVolHeadroom is used.

' All data used by this method is passed in as arguments (with the exception of look-ups into the Lines workbook)
' The method updates the PFE sheet and populates a dictionary D
'
'D is populated with items with the following names
' NumTrades           The number of trades valued
' FxShock             The Fx shock in use (does not affect the result for Mode = SolveFx)
' FxVolShock          The Fx vol shock in use
' MinHeadroomOverFirstN  Five element array. Element n is the minimum over the first n years of "Line - PFE" or
'                     "Line - Notional summation"
' MaxPFEByYear        Five element array. Element n is the maximum  from t = n-1 to t = n of PFE
' TradeSolveResult    A string either "OK" or an error of some kind...
' TradeHeadroom       Five element array. Element n is the max EURUSD trade that can be done (in isolation) to exhaust
'                     the n-year line. trade is measured in USD
' TradeHeadroom345    Five element array (first 2 of which always zero) giving the the maximum notional that could be
'                     executed simultaneously in 3,4, and 5 year EURUSD forward trades
' FxSolveResult       A string either "OK" or an error of some kind...
' FxHeadroomCol1      Five element array. Element n is the level of EURUSD spot that exhausts the n-year line.
'                     For HW model all these are the same and are the
' -----------------------------------------------------------------------------------------------------------------------
Private Function RunCreditUsageSheetCore(Mode As String, FilterBy1 As String, Filter1Value, FilterBy2 As String, Filter2Value, _
          IncludeFutureTrades As Boolean, ByVal IncludeAssetClasses As String, _
          IncludeExtraTrades As Boolean, ExtraTradeLabels, ExtraTradeAmounts, PortfolioAgeing As Double, _
          TradesScaleFactor As Double, CurrenciesToInclude As String, NumMCPaths As Long, _
          ByVal NumObservations As Double, FxShock As Double, FxVolShock As Double, LinesScaleFactor As Double, _
          ThrowErrors As Boolean, DisplayBaseCasePFE As Boolean, RefreshSheet As Boolean, _
          ByRef ResultsDict As Dictionary, Allocations As Variant, ExtraTradesAre As String)
        
          Const NumMCPathsError = "NumMCPaths must be one less than a power of 2 (e.g. 127, 255 or 1023)"
          Dim AnchorDate As Date
          Dim BaseCCY As String
          Dim CalcByProduct As Boolean
          Dim ChartTitle As String
          Dim CLFPLeftCol
          Dim CLFPRightCol
          Dim CreditLimits As Variant
          Dim CreditLimitsForPlotting As Variant
          Dim CreditLineInterp As String
          Dim DoNotionalCap As Boolean
          Dim EitherOrBasis As Boolean
          Dim ErrorMessage As String
          Dim FlipTrades As Boolean
          Dim FxNotionalPercentages As Variant
          Dim FxSolveResult As String
          Dim FxVolSolveResult As String
          Dim Headers As Variant
          Dim Headroom As Variant
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim InterpolatedLines
          Dim IsNB As Boolean
          Dim MaxPFEByYear As Variant
          Dim Methodology As String
          Dim MinHeadroomOverFirstN As Variant
          Dim ModelName As String
          Dim Multiple As Variant
          Dim NotionalCap As Variant
          Dim NotionalCapApplies As Boolean
          Dim Numeraire As String
          Dim oldBlockChange As Boolean
          Dim PFEPercentile As Variant
          Dim PFEProfile
          Dim PFEProfileShockedFx
          Dim PFEProfileUnshockedFx
          Dim PFEProfileWithET
          Dim PFEProfileWithoutET
          Dim PFEToDisplay
          Dim PFEVector As Variant
          Dim Pow2 As Long
          Dim ProductCreditLimits As String
          Dim ProfileResult As Variant
          Dim RatesNotionalPercentages As Variant
          Dim ShockedModel As Dictionary
          Dim ShortfallOrQuantile As String
          Dim SolveResult As Variant
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TC As TradeCount
          Dim TC2 As TradeCount
          Dim TimeEnd As Double
          Dim TimeGap As Double
          Dim TimeVector As Variant
          Dim TradeHeadroom345 As Variant
          Dim TradeSolveResult As String
          Dim twb As Workbook
          Dim UnitHedgeAmounts As Variant
          Dim UseHistoricalFxVol As Boolean
          Dim VolatilityInput As String
          Dim xArrayAscending As Variant
          Dim yArray As Variant
          Dim YAxisTitle As String

1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent: gBlockChangeEvent = True

          'Get information from the Lines workbook and do early error checking on inputs
3         Methodology = FirstElement(LookupCounterpartyInfo(Filter1Value, "Methodology", " - ", " - "))
4         IsNB = LCase(Methodology) = "notional based"

5         ClearoutResults

6         VolatilityInput = IIf(IsNB, "-", FirstElement(LookupCounterpartyInfo(Filter1Value, "Volatility Input", "MARKET IMPLIED", "MARKET IMPLIED")))
7         CreditLineInterp = FirstElement(LookupCounterpartyInfo(Filter1Value, "Line Interp.", "FlatToRight", " - ")) 'N.B. that " - " is tested for in method ValidateCreditLimits
8         ProductCreditLimits = FirstElement(LookupCounterpartyInfo(Filter1Value, "Product Credit Limits", "Global Calculation", "Global Calculation"))
9         NotionalCap = FirstElement(LookupCounterpartyInfo(Filter1Value, "Notional Cap", " - ", " - "))
10        DoNotionalCap = IsNumber(NotionalCap)
11        ProductCreditLimits = Replace(ProductCreditLimits, "Calculation", "Calc")
12        SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, IncludeExtraTrades, IncludeFxTrades, IncludeRatesTrades, CalcByProduct

13        PFEPercentile = IIf(IsNB, "-", FirstElement(LookupCounterpartyInfo(Filter1Value, "Confidence %", 0.95, 0.95)))
14        ShortfallOrQuantile = IIf(IsNB, "-", FirstElement(LookupCounterpartyInfo(Filter1Value, "Shortfall or Quantile", "-", "Quantile")))
15        BaseCCY = UCase(Left(Trim(FirstElement(LookupCounterpartyInfo(Filter1Value, "Base Currency", "EUR", "EUR"))), 3))
16        If InStr(LCase(CurrenciesToInclude), LCase(BaseCCY)) = 0 Then Throw "The Bank's 'Base Currency' (currently " & BaseCCY & ") must be listed in CurrenciesToInclude on the Config sheet"
17        If IsNB Then
18            FxNotionalPercentages = UnPackNotionalPercentages(FirstElementOf(LookupCounterpartyInfo(Filter1Value, "Fx Notional Weights")), False)
19            RatesNotionalPercentages = UnPackNotionalPercentages(FirstElementOf(LookupCounterpartyInfo(Filter1Value, "Rates Notional Weights")), True)
20        Else
21            FxNotionalPercentages = "-"
22            RatesNotionalPercentages = "-"
23        End If

24        UseHistoricalFxVol = VolatilityInput = "HISTORICAL"
25        FlipTrades = True 'So that the trades are represented from the Banks' point of view

26        If NumMCPaths < 1 Then
27            Throw NumMCPathsError
28        Else
29            Pow2 = Log(NumMCPaths + 1) / Log(2)
30            If 2 ^ Pow2 - 1 <> NumMCPaths Then Throw NumMCPathsError
31        End If

32        If NumObservations < 1 Or NumObservations > 2000 Or CLng(NumObservations) <> NumObservations Then Throw "NumObservations must be a whole number in the range 1 to 2000"
33        TimeEnd = GetHedgeHorizon()

34        TimeGap = TimeEnd / NumObservations
35        If FxShock <= 0 Then Throw "FxShock must be positive"

          'Prepare arguments for line interpolation.
36        If LinesScaleFactor <= 0 Then Throw "LinesScaleFactor must be positive"
37        xArrayAscending = sArrayStack(1, 2, 3, 4, 5, 7, 10)
          'Converting empty cells to 0 in call below
38        yArray = LookupCounterpartyInfo(Filter1Value, sArrayStack("1Y Limit", "2Y Limit", "3Y Limit", "4Y Limit", "5Y Limit", "7Y Limit", "10Y Limit"), 0, 0)
          Dim BankIsGood As Boolean
39        BankIsGood = FirstElementOf(LookupCounterpartyInfo(Filter1Value, "CPTY_PARENT", , "Bank not found")) = Filter1Value

40        If BankIsGood Then
41            If VarType(yArray) = vbString Then
42                yArray = sReshape(0, 7, 1)
43            Else
                  Dim Tmp
44                For Each Tmp In yArray
45                    If Not IsNumber(Tmp) Then Throw "Credit Limits must be numbers"
46                Next
47                If LinesScaleFactor <> 1 Then yArray = sArrayMultiply(LinesScaleFactor, yArray)
48            End If
49        End If

50        If UseHistoricalFxVol Then
51            ModelName = MN_CMHS        'We don't "need" it i.e. look at it but we do need its name to be set!
52        Else
53            ModelName = MN_CMS
54        End If
           
55        BuildModelsInJulia False, FxShock, FxVolShock

56        If UseHistoricalFxVol Then
57            Set ShockedModel = gModel_CMHS
58            ModelName = MN_CMHS
59        Else
60            Set ShockedModel = gModel_CMS
61            ModelName = MN_CMS
62        End If

          Dim BaseEUR
          Dim BaseUSD
          Dim EURUSD
          Dim EURUSD3YVol
          Dim PVBase
          Dim PVEUR
          Dim PVUSD
          
63        AnchorDate = ShockedModel("AnchorDate")
64        BaseUSD = MyFxPerBaseCcy(BaseCCY, "USD", ShockedModel)
65        BaseEUR = MyFxPerBaseCcy(BaseCCY, "EUR", ShockedModel)

66        EURUSD = ShockedModel("EURUSD")
67        EURUSD3YVol = ShockedModel("EURUSD3YVol")

68        DictAdd ResultsDict, "EURUSD", EURUSD
69        DictAdd ResultsDict, "EURUSD3YVol", EURUSD3YVol

70        Numeraire = ShockedModel("Numeraire")
71        Set twb = OpenTradesWorkbook(True, False)

          'Handle Notional Cap
72        If DoNotionalCap Then
              Dim CurrentNotional
              Dim FakeFxNP
              Dim FakeRatesNP
              Dim NotionalCapForNewTrades As Double
              Dim Trades
73            Trades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
                  IncludeFutureTrades, PortfolioAgeing, FlipTrades, Numeraire, IncludeFxTrades, IncludeRatesTrades, _
                  TradesScaleFactor, CurrenciesToInclude, False, TC, twb, shFutureTrades, AnchorDate)
74            If TC.NumIncluded > 0 Then
75                FakeFxNP = sArraySquare(0, 0.01, 10, 0.01)
76                FakeRatesNP = sParseArrayString("{""Tenor"",""EUR"",""Other"";1,0.005,0.005;7,0.03,0.03}")
77                CurrentNotional = ThrowIfError(NotionalBasedFromTrades("USD", FakeFxNP, FakeRatesNP, ShockedModel, _
                      Trades, Empty, GetHedgeHorizon(), 1, Repeat(0, TC.NumIncluded), "TotalNotionalForCap"))
78            Else
79                CurrentNotional = 0
80            End If
81            NotionalCapForNewTrades = NotionalCap - CurrentNotional
82            If NotionalCapForNewTrades < 0 Then NotionalCapForNewTrades = 0
83            NotionalCapApplies = NotionalCap <= (CurrentNotional + IIf(IncludeExtraTrades, sColumnSum(sArrayAbs(ExtraTradeAmounts))(1, 1), 0))
84        Else
85            CurrentNotional = " - "
86        End If

          Const CompressTrades = True
          Dim Success As Boolean

87        Select Case Mode

              Case "Standard"
88                ProfileResult = PFEProfileFromFiltersPCL(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
                      IncludeAssetClasses, IncludeExtraTrades, ExtraTradeLabels, ExtraTradeAmounts, _
                      PortfolioAgeing, FlipTrades, TradesScaleFactor, BaseCCY, _
                      ModelName, NumMCPaths, TimeGap, TimeEnd, Methodology, PFEPercentile, _
                      ShortfallOrQuantile, FxNotionalPercentages, RatesNotionalPercentages, CurrenciesToInclude, _
                      ModelName, TC, ProductCreditLimits, twb, CompressTrades, ShockedModel, ExtraTradesAre)

89                If sIsErrorString(ProfileResult) Then
90                    ProfileResult = FirstElementOf(ProfileResult)
91                    If ThrowErrors Then Throw ProfileResult
92                Else
93                    Success = True
94                    PFEProfile = ProfileResult
95                    ProfileResult = "OK"
96                End If
97                DictAdd ResultsDict, "ProfileResult", ProfileResult

98                PFEToDisplay = PFEProfile
99                If IncludeExtraTrades Then
100                   PFEProfileWithET = PFEProfile
101               Else
102                   PFEProfileWithoutET = PFEProfile
103               End If

104           Case "Solve345"
105               CheckCreditLimits FilterBy1, yArray
106               If sNRows(Allocations) <> GetHedgeHorizon() Then
107                   Throw "Allocations should have " & CStr(GetHedgeHorizon()) & " rows, but got " & CStr(sNRows(Allocations))
108               End If
109               UnitHedgeAmounts = sArrayMultiply(100000000#, Allocations)

110               If IsNB Then

                      'Line below sets ByRef arguments PFEProfile and NumTrades
111                   SolveResult = NBSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, IncludeAssetClasses, ExtraTradeLabels, _
                          UnitHedgeAmounts, PortfolioAgeing, FlipTrades, ModelName, TimeGap, _
                          TimeEnd, FxNotionalPercentages, RatesNotionalPercentages, xArrayAscending, yArray, CreditLineInterp, _
                          BaseCCY, TradesScaleFactor, PFEProfileWithET, PFEProfileWithoutET, CurrenciesToInclude, TC, ProductCreditLimits, _
                          DoNotionalCap, NotionalCapApplies, NotionalCapForNewTrades, twb, shFutureTrades, ShockedModel, ExtraTradesAre)

112                   If sIsErrorString(SolveResult) Then
113                       Multiple = 0
114                       TradeSolveResult = FirstElementOf(SolveResult)
115                       If ThrowErrors Then Throw CStr(TradeSolveResult)
116                   Else
117                       Success = True
118                       Multiple = FirstElementOf(SolveResult)
119                       TradeHeadroom345 = sArrayMultiply(Multiple, UnitHedgeAmounts)
120                       TradeSolveResult = IIf(NotionalCapApplies, "OK, Notional Cap Applies", "OK")
121                   End If
122                   PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileWithoutET, PFEProfileWithET)
123               Else
124                   EitherOrBasis = False
                      'Line below sets ByRef arguments PFEProfile and NumTrades
125                   SolveResult = HeadroomSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
                          IncludeAssetClasses, ExtraTradeLabels, UnitHedgeAmounts, PortfolioAgeing, _
                          FlipTrades, Numeraire, NumMCPaths, TimeGap, TimeEnd, CDbl(PFEPercentile), ShortfallOrQuantile, _
                          xArrayAscending, yArray, CreditLineInterp, BaseCCY, TradesScaleFactor, _
                          PFEProfileWithET, PFEProfileWithoutET, CurrenciesToInclude, ModelName, TC, ProductCreditLimits, _
                          DoNotionalCap, NotionalCapApplies, NotionalCapForNewTrades, twb, ShockedModel, EitherOrBasis, ExtraTradesAre)
126                   If sIsErrorString(SolveResult) Then
127                       Multiple = 0
128                       TradeSolveResult = FirstElementOf(SolveResult)
129                       If ThrowErrors Then Throw SolveResult
130                   Else
131                       Success = True
132                       Multiple = FirstElementOf(SolveResult)
133                       TradeHeadroom345 = sArrayMultiply(Multiple, UnitHedgeAmounts)
134                       If DoNotionalCap Then
135                           If sColumnSum(sArrayAbs(TradeHeadroom345))(1, 1) > NotionalCapForNewTrades Then
136                               NotionalCapApplies = True
137                               TradeSolveResult = "OK, Notional Cap Applies"
138                               TradeHeadroom345 = sArrayMultiply(TradeHeadroom345, NotionalCapForNewTrades / sColumnSum(sArrayAbs(TradeHeadroom345))(1, 1))
139                           End If
140                       End If
141                       TradeSolveResult = "OK"
142                   End If

143                   PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileWithoutET, PFEProfileWithET)
144               End If
145           Case "Solve1to5"
                  Dim TradeHeadroom As Variant
146               CheckCreditLimits FilterBy1, yArray
147               If IsNB Then
148                   SolveResult = NBSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, IncludeAssetClasses, ExtraTradeLabels, _
                          sArrayMultiply(1000000, sIdentityMatrix(sNRows(ExtraTradeLabels))), PortfolioAgeing, _
                          FlipTrades, ModelName, TimeGap, TimeEnd, FxNotionalPercentages, RatesNotionalPercentages, _
                          xArrayAscending, yArray, CreditLineInterp, BaseCCY, TradesScaleFactor, PFEProfileWithET, PFEProfileWithoutET, _
                          CurrenciesToInclude, TC, ProductCreditLimits, DoNotionalCap, NotionalCapApplies, NotionalCapForNewTrades, twb, shFutureTrades, ShockedModel, ExtraTradesAre)

149                   If sIsErrorString(SolveResult) Then
150                       TradeHeadroom = sReshape(0, GetHedgeHorizon(), 1)
151                       If ThrowErrors Then Throw SolveResult
152                       TradeSolveResult = FirstElementOf(SolveResult)
153                   Else
154                       Success = True
155                       TradeSolveResult = "OK"
156                       TradeHeadroom = sArrayMultiply(SolveResult, 1000000)
157                       If NotionalCapApplies Then TradeSolveResult = "OK, Notional Cap Applies"
158                       PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileWithoutET, PFEProfileWithET)
159                   End If
160               Else
161                   EitherOrBasis = True
162                   TradeHeadroom = sReshape(0, GetHedgeHorizon(), 1)
163                   TradeSolveResult = "OK"        'Gets overwritten by an error string if an error happens
164                   UnitHedgeAmounts = sReshape(100000000#, GetHedgeHorizon(), 1)
165                   SolveResult = HeadroomSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
                          IncludeAssetClasses, ExtraTradeLabels, UnitHedgeAmounts, PortfolioAgeing, _
                          FlipTrades, Numeraire, NumMCPaths, TimeGap, TimeEnd, CDbl(PFEPercentile), ShortfallOrQuantile, _
                          xArrayAscending, yArray, CreditLineInterp, BaseCCY, TradesScaleFactor, _
                          PFEProfileWithET, PFEProfileWithoutET, CurrenciesToInclude, ModelName, TC, ProductCreditLimits, _
                          DoNotionalCap, NotionalCapApplies, NotionalCapForNewTrades, twb, ShockedModel, EitherOrBasis, ExtraTradesAre)

166                   If sIsErrorString(SolveResult) Then
167                       Multiple = 0
168                       If ThrowErrors Then Throw SolveResult
169                       TradeHeadroom = sReshape(SolveResult, 5, 1)
170                   Else
171                       Success = True
172                       TradeHeadroom = sArrayMultiply(sArrayTranspose(SolveResult), 100000000#)
173                   End If
174                   If DoNotionalCap Then
175                       If sMaxOfArray(TradeHeadroom) >= NotionalCapForNewTrades Then
176                           TradeHeadroom = sArrayMin(TradeHeadroom, NotionalCapForNewTrades)
177                           NotionalCapApplies = True 'At least for one of the five
178                           TradeSolveResult = "Ok, Notional Cap Applies"
179                       End If
180                   End If

181                   PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileWithoutET, PFEProfileWithET)
182               End If

183           Case "SolveFx"
184               CheckCreditLimits FilterBy1, yArray
185               If IsNB Then
186                   SolveFxHeadroom ThrowErrors, DisplayBaseCasePFE, Success
187                   GoTo CollectResultsFromSheet
188               End If
                  Dim FxRoot As Variant
189               SolveResult = FxSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
                      IncludeAssetClasses, PortfolioAgeing, FlipTrades, _
                      Numeraire, NumMCPaths, TimeGap, TimeEnd, CDbl(PFEPercentile), ShortfallOrQuantile, _
                      xArrayAscending, yArray, CreditLineInterp, BaseCCY, TradesScaleFactor, _
                      PFEProfileUnshockedFx, PFEProfileShockedFx, CurrenciesToInclude, ModelName, TC, CalcByProduct, _
                      ProductCreditLimits, twb, FxRoot, ShockedModel)

190               If sIsErrorString(SolveResult) Then
191                   Multiple = 0
192                   FxSolveResult = SolveResult
193                   If ThrowErrors Then Throw SolveResult
194               Else
195                   Success = True
196                   Multiple = SolveResult * FxShock
197                   FxSolveResult = "OK"
198               End If

199               PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileUnshockedFx, PFEProfileShockedFx)

200           Case "SolveFxVol"
201               CheckCreditLimits FilterBy1, yArray
202               If IsNB Then
203                   Throw "Solving for FxVol headroom is not supported for banks which use a notional based approach, since FxVol is not an explicit input to the line-use calculation"
204               End If
                  Dim FxVolRoot ' (set by reference in call to FxVolSolverFromFilters)
205               SolveResult = FxVolSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
                      IncludeAssetClasses, PortfolioAgeing, FlipTrades, _
                      Numeraire, NumMCPaths, TimeGap, TimeEnd, CDbl(PFEPercentile), ShortfallOrQuantile, _
                      xArrayAscending, yArray, CreditLineInterp, BaseCCY, TradesScaleFactor, _
                      PFEProfileUnshockedFx, PFEProfileShockedFx, CurrenciesToInclude, ModelName, TC, CalcByProduct, _
                      ProductCreditLimits, twb, FxVolRoot, ShockedModel)

206               If sIsErrorString(SolveResult) Then
207                   Multiple = 0
208                   FxVolSolveResult = SolveResult
                      ' FxVolHeadroomCol1 = sReshape(0, 5, 1)
209                   If ThrowErrors Then Throw SolveResult
210               Else
211                   Success = True
212                   Multiple = SolveResult * FxVolShock
                      ' FxVolHeadroomCol1 = sReshape(FxVolRoot, 5, 1)
213                   FxVolSolveResult = "OK"
214               End If

215               PFEToDisplay = IIf(DisplayBaseCasePFE, PFEProfileUnshockedFx, PFEProfileShockedFx)

216           Case Else
217               Throw "Unrecognised Mode. Recognised values are: Standard, Solve345 and Solve1to5, SolveFx, SolveFxVol"
218       End Select

          'Calculate data to go into D
219       TimeVector = sSubArray(PFEToDisplay, 1, 2, , 1)
220       PFEVector = sSubArray(PFEToDisplay, 1, 3, , 1)
221       If BankIsGood Then
222           InterpolatedLines = ThrowIfError(sInterp(xArrayAscending, yArray, TimeVector, CreditLineInterp, "FF"))
223           CreditLimits = sArrayRange(xArrayAscending, yArray)
224           Headroom = sArraySubtract(InterpolatedLines, PFEVector)
225           MaxPFEByYear = MaxByBuckets(GetHedgeHorizon(), TimeVector, PFEVector)
226           MinHeadroomOverFirstN = MinOverFirstN(GetHedgeHorizon(), TimeVector, InterpolatedLines, PFEVector)
227       Else
228           InterpolatedLines = sReshape(Empty, sNRows(PFEToDisplay), 1)
229           Headroom = InterpolatedLines
230           MaxPFEByYear = MaxByBuckets(GetHedgeHorizon(), TimeVector, PFEVector)
231           MinHeadroomOverFirstN = sReshape(Empty, GetHedgeHorizon(), 1)
232       End If

          'FxShocks are different from when the variables BaseUSD, BaseEUR, EURUSD were populated so re-populate...
233       If Mode = "SolveFx" Then
234           If Not sIsErrorString(SolveResult) Then
235               If BaseCCY = "EUR" Then
236                   BaseUSD = MyFxPerBaseCcy(BaseCCY, "USD", gModel_CM) * Multiple
237               End If
238               If BaseCCY <> "EUR" Then
239                   BaseEUR = MyFxPerBaseCcy(BaseCCY, "EUR", gModel_CM) / Multiple
240               End If
241               EURUSD = GetItem(gModel_CM, "EURUSD") * Multiple
242           End If
243       End If

244       If Not IsNB Then
245           If sNCols(PFEToDisplay) >= 3 Then
246               PVBase = PFEToDisplay(1, 3)        'No point in calculating again, but we have to be careful that the PV we extract from the t=0 PFE graph _
                                                      is correct - e.g. does not include ExtraTrades (though such trades are meant to have zero PV).

247               PVEUR = PVBase * BaseEUR
248               PVUSD = PVBase * BaseUSD
249           ElseIf sIsErrorString(PFEToDisplay) Then
250               PVBase = FirstElementOf(PFEToDisplay): PVEUR = PVBase: PVUSD = PVBase
251           Else
252               PVBase = "#Unknown error!": PVEUR = PVBase: PVUSD = PVBase
253           End If
254       Else
255           PVBase = FirstElementOf(PortfolioValueFromFilters(BaseCCY, FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, IncludeAssetClasses, _
                  PortfolioAgeing, TradesScaleFactor, CurrenciesToInclude, ModelName, TC2, ProductCreditLimits, twb, AnchorDate, True))
256           If IsNumber(PVBase) Then
257               PVEUR = PVBase * BaseEUR
258               PVUSD = PVBase * BaseUSD
259           Else
260               PVEUR = PVBase
261               PVUSD = PVBase
262           End If
263       End If

264       DictAdd ResultsDict, "Success", Success
265       DictAdd ResultsDict, "PVBase", PVBase
266       DictAdd ResultsDict, "PVEUR", PVEUR
267       DictAdd ResultsDict, "PVUSD", PVUSD
268       DictAdd ResultsDict, "NumTrades", TC.NumIncluded
269       DictAdd ResultsDict, "NumTradesExcluded", TC.NumExcluded
270       DictAdd ResultsDict, "FxShock", FxShock
271       DictAdd ResultsDict, "FxVolShock", FxVolShock
272       DictAdd ResultsDict, "MinHeadroomOverFirstN", MinHeadroomOverFirstN
273       DictAdd ResultsDict, "MaxPFEByYear", MaxPFEByYear
274       If Mode = "Solve1to5" Then
275           DictAdd ResultsDict, "TradeSolveResult", TradeSolveResult
276           DictAdd ResultsDict, "TradeHeadroom", TradeHeadroom
277       ElseIf Mode = "Solve345" Then
278           DictAdd ResultsDict, "TradeSolveResult", TradeSolveResult
279           DictAdd ResultsDict, "TradeHeadroom345", TradeHeadroom345
280       ElseIf Mode = "SolveFx" Then
281           DictAdd ResultsDict, "FxSolveResult", FxSolveResult
282           DictAdd ResultsDict, "FxRoot", FxRoot
283       ElseIf Mode = "SolveFxVol" Then
284           DictAdd ResultsDict, "FxVolSolveResult", FxVolSolveResult
285           DictAdd ResultsDict, "FxVolRoot", FxVolRoot
286       End If
287       DictAdd ResultsDict, "MinHeadroomOverFirstNUSD", sArrayMultiply(MinHeadroomOverFirstN, BaseUSD)

288       If Not RefreshSheet Then GoTo EarlyExit

          'Stuff needed only for chart
289       If CreditLineInterp = "FlatToRight" Then
290           CLFPLeftCol = sArrayStack(0, sGroupReshape(xArrayAscending, 2))
291           CLFPRightCol = sArrayStack(sGroupReshape(yArray, 2), 0)
292       Else
293           CLFPLeftCol = sArrayStack(0, xArrayAscending)
294           CLFPRightCol = sArrayStack(yArray(1, 1), yArray)
295       End If

296       If DoNotionalCap Then
297           Select Case Mode
                  Case "Solve345"
298                   CurrentNotional = CurrentNotional + FirstElementOf(sColumnSum(sArrayAbs(TradeHeadroom345)))
299                   If CurrentNotional >= NotionalCap Then NotionalCapApplies = True
300               Case "Standard", "SolveFx"
301                   If IncludeExtraTrades Then
302                       CurrentNotional = CurrentNotional + FirstElementOf(sColumnSum(sArrayAbs(ExtraTradeAmounts)))
303                       If CurrentNotional >= NotionalCap Then NotionalCapApplies = True
304                   End If

305               Case "Solve1to5"
306                   CurrentNotional = CurrentNotional + Abs(TradeHeadroom(GetHedgeHorizon(), 1))        'we leave the sheet in the state as it would be after doing the last headroom calculation
307                   If CurrentNotional >= NotionalCap Then NotionalCapApplies = True
308           End Select
309       End If

310       CreditLimitsForPlotting = sArrayRange(CLFPLeftCol, CLFPRightCol)

          Dim PorS As String
311       If LCase(ShortfallOrQuantile) = "shortfall" Then
312           PorS = "ES"
313       Else
314           PorS = "PFE"
315       End If

316       Headers = sArrayRange("Date", "Time", _
              IIf(IsNB, "Notional-based line usage", PorS & _
              "(" & Format(PFEPercentile, "0%") & ")"), _
              "Line", IIf(IsNB, "Line Headroom", "Headroom"))

          'We've calculated everything, so update the sheet...
317       Set SPH = CreateSheetProtectionHandler(shCreditUsage)
318       Set SUH = CreateScreenUpdateHandler()
319       If TypeName(Selection) <> "Range" Then
              'Horrible possibility that user has selected the chart while the macro is running - possible because of the DoEvents call in RefreshScreen?
320           shCreditUsage.Cells(1, 1).Select
321       End If

          Dim DataToPaste As Variant
322       DataToPaste = sArrayRange(PFEToDisplay, InterpolatedLines, Headroom)

          Dim DifNumRows As Boolean
323       With RangeFromSheet(shCreditUsage, "TheData")
324           DifNumRows = .Rows.Count <> sNRows(DataToPaste)
325           .Rows(0).Value = Headers
326           If DifNumRows Then
327               .Clear
328           Else
329               .ClearContents
330           End If
331           With .Resize(sNRows(PFEToDisplay))
332               .Value = DataToPaste
333               If DifNumRows Then
334                   shCreditUsage.Names.Add "TheData", .offset(0)
335               End If
336           End With
337       End With

338       RangeFromSheet(shCreditUsage, "MaxPFEByYear").Value = MaxPFEByYear
339       RangeFromSheet(shCreditUsage, "MinHeadroomOverFirstN").Value = MinHeadroomOverFirstN
340       RangeFromSheet(shCreditUsage, "Methodology").Value = Methodology
341       RangeFromSheet(shCreditUsage, "PFEPercentile").Value = PFEPercentile
342       RangeFromSheet(shCreditUsage, "ShortfallOrQuantile").Value = ShortfallOrQuantile
343       RangeFromSheet(shCreditUsage, "VolatilityInput").Value = VolatilityInput
344       RangeFromSheet(shCreditUsage, "BaseCcy").Value = BaseCCY
345       RangeFromSheet(shCreditUsage, "CreditLineInterp").Value = CreditLineInterp
346       RangeFromSheet(shCreditUsage, "ProductCreditLimits").Value = ProductCreditLimits
347       RangeFromSheet(shCreditUsage, "CreditLimits").Value = CreditLimits
348       RangeFromSheet(shCreditUsage, "CreditLimitsForPlotting").Value = CreditLimitsForPlotting
349       RangeFromSheet(shCreditUsage, "FxNotionalPercentages").Value = FxNotionalPercentages
350       RangeFromSheet(shCreditUsage, "NotionalCap").Value = NotionalCap
351       RangeFromSheet(shCreditUsage, "CurrentNotional").Value = CurrentNotional
352       RangeFromSheet(shCreditUsage, "NumTrades").Value = TC.NumIncluded
353       RangeFromSheet(shCreditUsage, "NumTradesValued").Value = TC.NumIncluded
354       RangeFromSheet(shCreditUsage, "NumTradesExcluded").Value = TC.NumExcluded
355       RangeFromSheet(shCreditUsage, "EURUSD").Value = EURUSD
356       RangeFromSheet(shCreditUsage, "EURUSD3YVol").Value = EURUSD3YVol
357       RangeFromSheet(shCreditUsage, "PVUSD").Value = PVUSD
358       RangeFromSheet(shCreditUsage, "PVEUR").Value = PVEUR

359       Select Case Mode
              Case "Solve345"
360               RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value = True
361               RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value = TradeHeadroom345
362               RunCreditUsageSheetCore = TradeHeadroom345
363               GreyOutHeadrooms True, True, True
364           Case "Solve1to5"
365               RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value = True
366               RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value = 0
367               RangeFromSheet(shCreditUsage, "TradeHeadroom").Value = TradeHeadroom
368               GreyOutHeadrooms False, True, True
369               RunCreditUsageSheetCore = "OK"
370           Case "Standard"
371               GreyOutHeadrooms True, True, True
372               RunCreditUsageSheetCore = "OK"
373           Case "SolveFx"
374               RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value = False
375               RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 2).Value = Multiple
376               RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 1).Value = FxRoot
377               RangeFromSheet(shCreditUsage, "FxShock").Value = Multiple
378               GreyOutHeadrooms True, False, True
379           Case "SolveFxVol"
380               RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value = False
381               RangeFromSheet(shCreditUsage, "FxVolHeadroom").Cells(1, 2).Value = Multiple
382               RangeFromSheet(shCreditUsage, "FxVolHeadroom").Cells(1, 1).Value = FxVolRoot
383               RangeFromSheet(shCreditUsage, "FxVolShock").Value = Multiple
384               GreyOutHeadrooms True, True, False
385       End Select

386       ChartTitle = PFEChartTitle(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeExtraTrades, ExtraTradeAmounts, _
              PortfolioAgeing, IIf(Mode = "SolveFx", FxRoot, FxShock), FxVolShock, TradesScaleFactor, _
              LinesScaleFactor, TC.NumIncluded, BankIsGood, IncludeFxTrades, IncludeRatesTrades, _
              IIf(NotionalCapApplies, "Notional Cap Applies", ""))
387       YAxisTitle = BaseCCY & " Millions"

388       UpdateChartOnCreditUsageSheet ChartTitle, YAxisTitle, BankIsGood

389       GoTo EarlyExit
          'NB this block of code must be in line with previous block where we populate the results collection
CollectResultsFromSheet:        'This is the case where we have had to use one of the "old" methods of solving (for LnFx model) )that pre-dates our _
                                       use of HullWhite, R and R's powerful root-finders. In these cases we don't have the option of not updating the PFE sheet

390       DictAdd ResultsDict, "Success", Success
391       PVBase = FirstElementOf(PortfolioValueFromFilters(BaseCCY, FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, IncludeAssetClasses, PortfolioAgeing, TradesScaleFactor, CurrenciesToInclude, _
              ModelName, TC2, ProductCreditLimits, twb, AnchorDate, True))
              
392       If IsNumber(PVBase) Then
393           PVEUR = PVBase * BaseEUR
394           PVUSD = PVBase * BaseUSD
395       Else
396           PVEUR = PVBase
397           PVUSD = PVBase
398       End If
399       DictAdd ResultsDict, "PVBase", PVBase
400       DictAdd ResultsDict, "PVEUR", PVEUR
401       DictAdd ResultsDict, "PVUSD", PVUSD
402       DictAdd ResultsDict, "NumTrades", RangeFromSheet(shCreditUsage, "NumTrades").Value
403       DictAdd ResultsDict, "FxShock", RangeFromSheet(shCreditUsage, "FxShock").Value
404       DictAdd ResultsDict, "FxVolShock", RangeFromSheet(shCreditUsage, "FxVolShock").Value
405       DictAdd ResultsDict, "MinHeadroomOverFirstN", RangeFromSheet(shCreditUsage, "MinHeadroomOverFirstN").Value
406       DictAdd ResultsDict, "MinHeadroomOverFirstNUSD", sArrayMultiply(ResultsDict("MinHeadroomOverFirstN"), BaseUSD)
407       DictAdd ResultsDict, "MaxPFEByYear", RangeFromSheet(shCreditUsage, "MaxPFEByYear").Value
408       If Mode = "Solve1to5" Then
409           DictAdd ResultsDict, "TradeSolveResult", RangeFromSheet(shCreditUsage, "TradeSolveResult").Value
410           DictAdd ResultsDict, "TradeHeadroom", RangeFromSheet(shCreditUsage, "TradeHeadroom").Value
411       ElseIf Mode = "Solve345" Then
412           DictAdd ResultsDict, "TradeSolveResult", RangeFromSheet(shCreditUsage, "TradeSolveResult").Value
413           DictAdd ResultsDict, "TradeHeadroom345", RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value        'This statement looks strange, but that's where the method dumps the data
414       ElseIf Mode = "SolveFx" Then
415           DictAdd ResultsDict, "FxSolveResult", RangeFromSheet(shCreditUsage, "FxSolveResult").Value
416           DictAdd ResultsDict, "FxRoot", RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 1).Value
417       End If

EarlyExit:
418       Set twb = Nothing
419       gBlockChangeEvent = oldBlockChange
420       Exit Function
ErrHandler:

421       ErrorMessage = "#RunCreditUsageSheetCore (line " & CStr(Erl) & "): " & Err.Description & "!"
422       gBlockChangeEvent = oldBlockChange
423       Set twb = Nothing
424       If ThrowErrors Then Throw ErrorMessage
425       DictAdd ResultsDict, "Success", False
426       Select Case Mode
              Case "Solve1to5", "Solve345"
427               DictAdd ResultsDict, "TradeSolveResult", ErrorMessage
428           Case "SolveFx"
429               DictAdd ResultsDict, "FxSolveResult", ErrorMessage
430           Case Else
431               DictAdd ResultsDict, "ProfileResult", ErrorMessage
432       End Select
433       RunCreditUsageSheetCore = ErrorMessage
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetAllocation
' Author    : Philip Swannell
' Date      : 21-Oct-2016
' Purpose   : Get input from the user for "custom allocation" trade headroom solving.
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetAllocation()
          Static PreviousResult As String
          Dim Allocation As Variant
          Dim HH As Long
          Dim LastError As String
          Dim NR As Long
          Dim Prompt As String
          Dim PromptFirstPart As String
          Dim Res As String

1         On Error GoTo ErrHandler
2         HH = GetHedgeHorizon()

3         PromptFirstPart = "Please enter trade allocation." & vbLf & vbLf & _
              "To allocate to a single year enter text such as '4y'." & vbLf & _
              "Otherwise, specify allocations with a colon-delimited" & vbLf & _
              "string. For example, to allocate equally to years 2, 4" & vbLf & _
              "and 5 enter '0:1:0:1:1'"

TryAgain:
4         If PreviousResult = "" Then
5             NR = 0
6         Else
7             NR = Len(PreviousResult) - Len(Replace(PreviousResult, ":", "")) + 1
8         End If

9         If NR > HH Then
10            PreviousResult = sConcatenateStrings(sTake(sTokeniseString(PreviousResult, ":"), HH), ":")
11        End If
12        If LastError = "" Then
13            Prompt = PromptFirstPart & vbLf & " "
14        Else
15            Prompt = PromptFirstPart & vbLf & vbLf & _
                  sConcatenateStrings(sJustifyText(LastError, "Calibri", 11, 300), vbLf) & "." & vbLf & " "
16        End If

17        Res = InputBoxPlus(Prompt, "Custom Trade Headroom", PreviousResult)
18        Res = Replace(Res, "'", "")
19        If Res = "False" Then
20            GetAllocation = Empty
21            Exit Function
22        End If

23        Allocation = Empty
24        On Error Resume Next
25        Allocation = ParseAllocation(Res, True)
26        LastError = Err.Description
27        On Error GoTo ErrHandler

28        If IsEmpty(Allocation) Then
29            Res = sConcatenateStrings(sReshape(1, HH, 1))
30            GoTo TryAgain
31        Else
32            GetAllocation = Allocation
33        End If
34        PreviousResult = Res
35        Exit Function
ErrHandler:
36        Throw "#GetAllocation (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PrepareForCalculation
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Make sure inputs are in a good state to run headrooms for one bank - e.g.
'             no portfolio ageing, no fx shock, no extra trades, filters correctly set etc.
' -----------------------------------------------------------------------------------------------------------------------
Sub PrepareForCalculation(BankName, ResetShocks As Boolean, ResetPortfolioAgeing As Boolean, ResetExtraTradesAre As Boolean)
          Dim CopyOfErr As String
          Dim oldBlockChange As Boolean
          Dim SPH As Object
1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent
3         gBlockChangeEvent = True

4         Set SPH = CreateSheetProtectionHandler(shCreditUsage)

5         If ResetPortfolioAgeing Then
6             With RangeFromSheet(shCreditUsage, "PortfolioAgeing")
7                 If .Value <> 0 Then .Value = 0
8             End With
9         End If

10        With RangeFromSheet(shCreditUsage, "IncludeExtraTrades")
11            If .Value <> False Then .Value = False
12        End With

13        With RangeFromSheet(shCreditUsage, "FilterBy1")
14            If .Value <> "Counterparty Parent" Then .Value = "Counterparty Parent"
15        End With

16        RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value = 0

17        If ResetShocks Or Not IsNumber(RangeFromSheet(shCreditUsage, "FxShock").Value) Then
18            With RangeFromSheet(shCreditUsage, "FxShock")
19                If .Value <> 1 Then .Value = 1
20            End With
21        End If
22        RangeFromSheet(shCreditUsage, "TradeHeadRoom").ClearContents
23        RangeFromSheet(shCreditUsage, "FxHeadroom").ClearContents
24        RangeFromSheet(shCreditUsage, "FxVolHeadroom").ClearContents
25        If ResetShocks Or Not IsNumber(RangeFromSheet(shCreditUsage, "FxVolShock").Value) Then
26            With RangeFromSheet(shCreditUsage, "FxVolShock")
27                If .Value <> 1 Then .Value = 1
28            End With
29        End If
30        With RangeFromSheet(shCreditUsage, "Filter1Value")
31            If .Value <> BankName Then .Value = BankName
32        End With

33        If ResetExtraTradesAre Then
34            RangeFromSheet(shCreditUsage, "ExtraTradesAre") = "Fx Airbus sells USD, buys EUR"
35        End If

36        gBlockChangeEvent = oldBlockChange

37        Exit Sub
ErrHandler:
38        CopyOfErr = "#PrepareForCalculation (line " & CStr(Erl) & "): " & Err.Description & "!"
39        gBlockChangeEvent = oldBlockChange
40        Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GreyOutHeadrooms
' Author    : Philip Swannell
' Date      : 03-Nov-2016
' Purpose   : Added this method after demo at Airbus 2 Nov 2016. It's sometimes unclear that the headroom
'             calculations are "stale" and might even relate to a different bank than the one now
'             held on the CreditUsage sheet.
' -----------------------------------------------------------------------------------------------------------------------
Sub GreyOutHeadrooms(GreyTradeHR As Boolean, GreyFxHR As Boolean, GreyFxVolHR As Boolean)

          Dim i As Long
          Dim isGrey As Boolean
          Dim RangeName As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shCreditUsage)
3         For i = 1 To 3
4             RangeName = Choose(i, "TradeHeadroom", "FxHeadroom", "FxVolHeadroom")
5             isGrey = Choose(i, GreyTradeHR, GreyFxHR, GreyFxVolHR)
6             With RangeFromSheet(shCreditUsage, RangeName)
7                 If isGrey Then
8                     .Font.Color = g_Col_GreyText
9                 Else
10                    .Font.ColorIndex = xlColorIndexAutomatic
11                End If
12            End With
13        Next i

14        Exit Sub
ErrHandler:
15        Throw "#GreyOutHeadrooms (line " & CStr(Erl) & "): " & Err.Description & "!"

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearoutResults
' Author    : Philip Swannell
' Date      : 04-Nov-2016
' Purpose   : Clear out results when we change which bank we are examining...
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearoutResults()
          Dim CheckSum As String
          Dim Filter1Value As String
          Dim Filter2Value As String
          Dim FilterBy1 As String
          Dim FilterBy2 As String
          Dim IncludeAssetClasses As String
          Dim RangeToClear As Range
          Static OldCheckSum As String
          Dim chtOb As ChartObject
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         FilterBy1 = CStr(RangeFromSheet(shCreditUsage, "FilterBy1").Value2)
3         Filter1Value = CStr(RangeFromSheet(shCreditUsage, "Filter1Value").Value2)
4         FilterBy2 = CStr(RangeFromSheet(shCreditUsage, "FilterBy2").Value2)
5         Filter2Value = CStr(RangeFromSheet(shCreditUsage, "Filter2Value").Value2)
6         IncludeAssetClasses = CStr(RangeFromSheet(shCreditUsage, "IncludeAssetClasses").Value2)

7         CheckSum = FilterBy1 & Filter1Value & FilterBy2 & Filter2Value & IncludeAssetClasses

8         If CheckSum <> OldCheckSum Then

9             Set SPH = CreateSheetProtectionHandler(shCreditUsage)
10            Set RangeToClear = Application.Union(Range(RangeFromSheet(shCreditUsage, "Methodology"), _
                  RangeFromSheet(shCreditUsage, "CurrentNotional")), _
                  Range(RangeFromSheet(shCreditUsage, "NumTradesValued"), _
                  RangeFromSheet(shCreditUsage, "NumTradesExcluded")), _
                  Range(RangeFromSheet(shCreditUsage, "PVUSD"), _
                  RangeFromSheet(shCreditUsage, "PVEUR")), _
                  Range(RangeFromSheet(shCreditUsage, "MaxPFEByYear"), _
                  RangeFromSheet(shCreditUsage, "FxVolHeadroom")), _
                  RangeFromSheet(shCreditUsage, "TheData"))

11            RangeToClear.ClearContents
12            For Each chtOb In shCreditUsage.ChartObjects
13                chtOb.Delete
14            Next
15        End If
16        OldCheckSum = CheckSum

17        Exit Sub
ErrHandler:
18        Throw "#ClearoutResults (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshScreen
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Update the screen and try to kick the Application status bar into life
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshScreen()
1         Application.ScreenUpdating = False
2         Application.ScreenUpdating = True
          '    DoEvents ' PGS 14 Nov 2016. Have hypothesis that DoEvents is causing hangs, so trying without
3         UnselectChart
4         Application.ScreenUpdating = False
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MaxByBuckets
' Author    : Philip Swannell
' Date      : 17-Jun-2015
' Purpose   : Utility function. Return is vector of height NumBuckets. ith element of return
'             is maximum of a) 0 and b) those elements of YVector for which corresponding element of
'             XVector, x is in i-1 < x <= i.
' -----------------------------------------------------------------------------------------------------------------------
Function MaxByBuckets(NumBuckets As Long, XVector, YVector)
          Dim i As Long
          Dim NR
          Dim Res() As Double
          Dim xRounded As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti XVector, YVector

3         ReDim Res(1 To NumBuckets, 1 To 1)

4         NR = sNRows(XVector)

5         For i = 1 To NR
6             If IsNumber(XVector(i, 1)) Then
7                 If IsNumber(YVector(i, 1)) Then
8                     If XVector(i, 1) = CLng(XVector(i, 1)) Then
9                         xRounded = XVector(i, 1)
10                    Else
11                        xRounded = CLng(XVector(i, 1) + 0.5)
12                    End If
13                    If xRounded > 0 Then
14                        If xRounded <= NumBuckets Then
15                            Res(xRounded, 1) = SafeMax(Res(xRounded, 1), YVector(i, 1))
16                        End If
17                    End If
18                End If
19            End If
20        Next i
21        MaxByBuckets = Res
22        Exit Function
ErrHandler:
23        Throw "#MaxByBuckets (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MinOverFirstN
' Author    : Philip Swannell
' Date      : 17-Jun-2015
' Purpose   : Utility function. Return is vector of height NumBuckets. ith element of return
'             is minimum of those elements of Headroom = (LinesVector-PFEVector) for which corresponding element of
'             TimeVector, x is in 0 <= x < i AND PFEVector <> 0
' -----------------------------------------------------------------------------------------------------------------------
Function MinOverFirstN(NumBuckets As Long, TimeVector, LinesVector, PFEVector)
          Dim i As Long
          Dim j As Long
          Dim NR
          Dim Res As Variant
          Dim TimeRounded As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti TimeVector, LinesVector, PFEVector

3         Res = sReshape(Empty, NumBuckets, 1)

4         NR = sNRows(TimeVector)

5         For i = 1 To NR
6             If IsNumber(TimeVector(i, 1)) Then
7                 If IsNumber(LinesVector(i, 1)) Then
8                     If IsNumber(PFEVector(i, 1)) Then
9                         If PFEVector(i, 1) <> 0 Then
10                            If TimeVector(i, 1) = CLng(TimeVector(i, 1)) Then
11                                TimeRounded = TimeVector(i, 1) + 1
12                            Else
13                                TimeRounded = CLng(TimeVector(i, 1) + 0.5)
14                            End If
15                            If TimeRounded > 0 Then
16                                If TimeRounded <= NumBuckets Then
17                                    For j = TimeRounded To NumBuckets
18                                        If IsEmpty(Res(j, 1)) Then
19                                            Res(j, 1) = LinesVector(i, 1) - PFEVector(i, 1)
20                                        Else
21                                            Res(j, 1) = SafeMin(Res(j, 1), LinesVector(i, 1) - PFEVector(i, 1))
22                                        End If
23                                    Next j
24                                End If
25                            End If
26                        End If
27                    End If
28                End If
29            End If
30        Next i

31        For i = 1 To NumBuckets
32            If IsEmpty(Res(i, 1)) Then
33                Res(i, 1) = "#No numbers found!"
34            End If
35        Next i

36        MinOverFirstN = Res
37        Exit Function
ErrHandler:
38        Throw "#MinOverFirstN (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function SafeEquals(a, b) As Boolean

1         On Error GoTo ErrHandler
2         SafeEquals = a = b
3         Exit Function
ErrHandler:
4         SafeEquals = False
End Function

'Calling SolumAddin.FormatAsInput is a bad idea - it's slow as it adds to the undo stack
'Call this method instead
Sub CayleyFormatAsInput(R As Range)
1         On Error GoTo ErrHandler
2         If Not SafeEquals(R.Font.Color, RGB(0, 0, 255)) Then
3             R.Font.Color = RGB(0, 0, 255)
4         End If
5         If Not SafeEquals(R.Locked, False) Then
6             R.Locked = False
7         End If
8         Exit Sub
ErrHandler:
9         Throw "#CayleyFormatAsInput (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub Test_StretchRanges()
1         StretchRanges GetHedgeHorizon()
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StretchRanges
' Author     : Philip Swannell
' Date       : 01-Feb-2022
' Purpose    :
' Parameters :
'  NewHeight:
' -----------------------------------------------------------------------------------------------------------------------
Private Sub StretchRanges(NewHeight As Long)

          Dim c As Range
          Dim i As Long
          Dim RangeName As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If RangeFromSheet(shCreditUsage, "ExtraTradeLabels").Rows.Count = GetHedgeHorizon() Then Exit Sub

3         Set SPH = CreateSheetProtectionHandler(shCreditUsage)

4         For i = 1 To 7
5             RangeName = Choose(i, "ExtraTradeLabels", "ExtraTradeAmounts", "MaxPFEByYear", "MinHeadroomOverFirstN", "TradeHeadroom", "FxHeadroom", "FxVolHeadroom")
6             StretchRange RangeName, NewHeight
7         Next i
        
8         With RangeFromSheet(shCreditUsage, "ExtraTradeLabels")
9             .Value = sArrayConcatenate(sIntegers(NewHeight), "Y")
10            AddGreyBorders .offset(-2).Resize(.Rows.Count + 2, 9), True
11        End With
12        For Each c In RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Cells
13            If Not IsNumber(c.Value) Then c.Value = 0
14        Next c

15        Exit Sub
ErrHandler:
16        Throw "#StretchRanges (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StretchRange
' Author     : Philip Swannell
' Date       : 01-Feb-2022
' Purpose    : Resize a range on CreditUsage worksheet to have HedgeHorizon() rows.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub StretchRange(RangeName As String, NewHeight As Long)

          Dim NR
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = RangeFromSheet(shCreditUsage, RangeName)
3         NR = R.Rows.Count

4         If NR = NewHeight Then Exit Sub

5         shCreditUsage.Names.Add RangeName, R.Resize(NewHeight)

6         If NR > GetHedgeHorizon() Then 'Shrinking ranges
7             RangeFromSheet(shCreditUsage, RangeName).offset(NewHeight).Resize(NR - NewHeight).Clear
8         End If

9         Exit Sub
ErrHandler:
10        Throw "#StretchRange (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetNumberFormat
' Author     : Philip Swannell
' Date       : 02-Mar-2022
' Purpose    : Set NumberFormat of a range optimised for speed if the range likely already has the correct number format
' -----------------------------------------------------------------------------------------------------------------------
Sub SetNumberFormat(R As Range, NumberFormat As String)
1         If Not SafeEquals(R.NumberFormat, NumberFormat) Then
2             R.NumberFormat = NumberFormat
3         End If
End Sub

Sub SetHorizontalAlignment(R As Range, HorizontalAlignment As Long)
1         If Not SafeEquals(R.HorizontalAlignment, HorizontalAlignment) Then
2             R.HorizontalAlignment = HorizontalAlignment
3         End If
End Sub

Sub SetLocked(R As Range, Locked As Boolean)
1         If Not SafeEquals(R.HorizontalAlignment, Locked) Then
2             R.Locked = Locked
3         End If
End Sub

Sub test_FormatCreditUsageSheet()
1         tic
2         FormatCreditUsageSheet False
3         toc "FormatCreditUsageSheet"

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatCreditUsageSheet
' Author    : Philip Swannell
' Date      : 07-Oct-2016
' Purpose   : Applies number formats etc to the PFE sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatCreditUsageSheet(ProtectAtExit As Boolean)
          Dim i As Long
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim ws As Worksheet
          
1         On Error GoTo ErrHandler
2         Set ws = shCreditUsage

3         Set SPH = CreateSheetProtectionHandler(ws)
4         Set SUH = CreateScreenUpdateHandler()
            
5         StretchRanges GetHedgeHorizon()

6         With RangeFromSheet(ws, "TheData")
7             If Not .Cells(1, 1).EntireColumn.Hidden Then
8                 SetHorizontalAlignment .Cells, xlHAlignCenter
9                 SetNumberFormat .Columns(1), "dd-mmm-yyyy"
10                SetNumberFormat .Columns(2), "0.000"
11                SetNumberFormat .Columns(3).Resize(, 3), "#,##0;[Red]-#,##0"
12                .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
13                AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
14            End If
15        End With

16        With RangeFromSheet(ws, "ExtraTradeLabels")
17            SetHorizontalAlignment .Cells, xlHAlignCenter
18            SetLocked .Cells, True
19        End With
20        With RangeFromSheet(ws, "MaxPFEByYear")
21            SetHorizontalAlignment .Cells, xlHAlignCenter
22            SetLocked .Cells, True
23            SetNumberFormat .Cells, "#,##0;[Red]-#,##0"
24        End With
25        With RangeFromSheet(ws, "MinHeadroomOverFirstN")
26            SetHorizontalAlignment .Cells, xlHAlignCenter
27            SetLocked .Cells, True
28            SetNumberFormat .Cells, "#,##0;[Red]-#,##0"
29        End With
30        With RangeFromSheet(ws, "TradeHeadroom")
31            SetHorizontalAlignment .Cells, xlHAlignCenter
32            SetLocked .Cells, True
33            SetNumberFormat .Cells, "#,##0;[Red]-#,##0"
34        End With

35        With RangeFromSheet(ws, "ExtraTradeAmounts")
36            .ClearFormats
37            SetNumberFormat .Cells, "#,##0;[Red]-#,##0"
38            SetHorizontalAlignment .Cells, xlHAlignCenter
39            SetLocked .Cells, False
40            CayleyFormatAsInput .Cells
41            .offset(-1, -1).Resize(.Rows.Count + 1, 9).Columns.AutoFit
42            For i = 0 To 3
43                If .offset(, i).ColumnWidth < 12.58 Then .offset(, i).ColumnWidth = 12.58
44            Next i
45            AddGreyBorders .offset(-2, -1).Resize(.Rows.Count + 2, 5), True
46            .FormatConditions.Delete
47            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & RangeFromSheet(ws, "IncludeExtraTrades").Address & "=FALSE"
48            .FormatConditions(.FormatConditions.Count).SetFirstPriority
49            With .FormatConditions(1).Font
50                .ThemeColor = xlThemeColorDark1
51                .TintAndShade = -0.249946592608417
52            End With
53            .FormatConditions(1).StopIfTrue = False
54        End With

55        With RangeFromSheet(ws, "FxHeadroom").Resize(, 4)
56            SetNumberFormat .Cells, "0.0000"
57            SetHorizontalAlignment .Cells, xlHAlignCenter
58            AddGreyBorders .offset(-2).Resize(3), True
59        End With

          Dim BottomName As String
          Dim TopName As String
60        For i = 1 To 4
61            TopName = Choose(i, "FilterBy1", "FxShock", "NumMCPaths", "TradesScaleFactor")
62            BottomName = Choose(i, "PortfolioAgeing", "FxVolShock", "NumObservations", "LinesScaleFactor")
63            With Range(RangeFromSheet(ws, TopName), RangeFromSheet(ws, BottomName))
64                SetLocked .Cells, False        'important in case user copies and pastes into a cell and thus locks it...
65                SetHorizontalAlignment .Cells, xlHAlignLeft
66                CayleyFormatAsInput .Cells
67            End With
68        Next i
69        SetNumberFormat RangeFromSheet(ws, "PortfolioAgeing"), "#,##0.000;[Red]-#,##0.000"

70        With RangeFromSheet(ws, "ModelType")
71            If .Value <> MT_HW Then
72                .Value = MT_HW
73            End If
74            SetLocked .Cells, True
75            .Font.ColorIndex = xlColorIndexAutomatic
76        End With

77        SetRangeValidation RangeFromSheet(ws, "IncludeAssetClasses"), sArrayStack("Rates and Fx", "Fx", "Rates"), True, _
              "Invalid IncludeAssetClasses", "IncludeAssetClasses"
78        SetRangeValidation RangeFromSheet(ws, "IncludeExtraTrades"), sArrayStack("TRUE", "FALSE"), True, _
              "Invalid IncludeExtraTrades", "IncludeExtraTrades"
79        SetRangeValidation RangeFromSheet(ws, "IncludeFutureTrades"), sArrayStack("TRUE", "FALSE"), True, _
              "Invalid IncludeFutureTrades", "IncludeFutureTrades"

80        If ProtectAtExit Then
81            Set SPH = Nothing
82            ws.Protect , True, True
83        End If

84        Exit Sub
ErrHandler:
85        Throw "#FormatCreditUsageSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub SetRangeValidation(R As Range, Choices As Variant, InCellDropdown As Boolean, ErrorMessage As String, ErrorTitle As String)
1         On Error GoTo ErrHandler
          Static ChoicesAsString As String
2         On Error Resume Next        ' It's bad to do this, but I seem to be getting very occasional errors and don't have time to track down why...

3         ChoicesAsString = sConcatenateStrings(Choices)

4         With R.Validation
5             .Delete
6             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=ChoicesAsString
7             .IgnoreBlank = True
8             .InCellDropdown = InCellDropdown
9             .InputTitle = ""
10            .ErrorTitle = ErrorTitle
11            .InputMessage = ""
12            .ErrorMessage = ErrorMessage
13            .ShowInput = True
14            .ShowError = True
15        End With

16        Exit Sub
ErrHandler:
17        Throw "#SetRangeValidation (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckCreditLimits
' Author    : Philip Swannell
' Date      : 29-Sep-2016
' Purpose   : Generate sensible error message when solving for headroom but with the filters
'             set to values that do not yield a credit line to solve against
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckCreditLimits(FilterBy1 As String, CreditLimits As Variant)
          Dim cl As Variant
1         If FilterBy1 <> "Counterparty Parent" Then Throw "FilterBy1 must be 'Counterparty Parent' to solve for headroom", True
2         For Each cl In CreditLimits
3             If Not (IsNumber(cl)) Then Throw "Error getting credit limits from Lines workbook - " & CStr(cl)
4         Next cl
5         On Error GoTo ErrHandler
6         Exit Function
ErrHandler:
7         Throw "#CheckCreditLimits (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddFilters
' Author    : Philip Swannell
' Date      : 21-Dec-2016
' Purpose   : Wrap to SCRiPTUtils.xlam!AddCayleyFiltersToMRU
' -----------------------------------------------------------------------------------------------------------------------
Sub AddFilters()
1         On Error GoTo ErrHandler
2         AddCayleyFiltersToMRU OpenTradesWorkbook(True, False), _
              RangeFromSheet(shCreditUsage, "FilterBy1", False, True, False, False, False).Value2, _
              RangeFromSheet(shCreditUsage, "Filter1Value", True, True, True, False, False).Value2, _
              RangeFromSheet(shCreditUsage, "FilterBy2", False, True, False, False, False).Value2, _
              RangeFromSheet(shCreditUsage, "Filter2Value", True, True, True, False, False).Value2, _
              Date
3         Exit Sub
ErrHandler:
4         Throw "#AddFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsItGettingPuffedOut
' Author    : Philip Swannell
' Date      : 28-Sep-2016
' Purpose   : Had suspicion that performance was degrading over time. But this method did not
'             provide any evidence of that
' -----------------------------------------------------------------------------------------------------------------------
Sub IsItGettingPuffedOut()
          Dim i As Long
1         For i = 1 To 100
2             MenuCreditUsageSheet "Calculate"
3         Next i
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeSetCellValue2
' Author    : Philip Swannell
' Date      : 09-Oct-2016
' Purpose   : Careful handling of setting cell values...
' -----------------------------------------------------------------------------------------------------------------------
Function SafeSetCellValue2(Target As Range, TheValue As String)
1         On Error GoTo ErrHandler
2         If IsNumeric(TheValue) Then
3             Target.Value = TheValue        'Excel will coerce to number
4         ElseIf UCase(TheValue) = "TRUE" Or UCase(TheValue) = "FALSE" Then
5             Target.Value = TheValue        'Excel will coerce to number'Excel will coerce to Boolean
6         Else
7             SafeSetCellValue Target, TheValue        'Avoid any Excel coercion with handling of first character being one of '@^\
8         End If
9         Exit Function
ErrHandler:
10        Throw "#SafeSetCellValue2 (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IfNotNumber
' Author    : Philip Swannell
' Date      : 30-Nov-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function IfNotNumber(TheData, ReplaceNonNumbersWith As Double)
1         IfNotNumber = sArrayIf(sArrayIsNumber(TheData), TheData, ReplaceNonNumbersWith)
End Function

Sub ShowHidePFEData(Optional Show)
          Const NarrowCaption = "Show data for chart"
          Const WideCaption = "Hide data for chart"
          Dim b As Button
          Dim Height As Double
          Dim Left As Double
          Dim LeftCell As Range
          Dim R As Range
          Dim RightCell As Range
          Dim SPH As clsSheetProtectionHandler
          Dim Top As Double
          Dim Width As Double

1         On Error GoTo ErrHandler
2         Set b = shCreditUsage.Buttons("butShowHideData")

3         If VarType(Show) <> vbBoolean Then
4             Show = b.Caption = NarrowCaption
5         End If

6         Set SPH = CreateSheetProtectionHandler(shCreditUsage)

7         Set LeftCell = RangeFromSheet(shCreditUsage, "TheData").Cells(1, 1)
8         Set RightCell = RangeFromSheet(shCreditUsage, "NumTrades")

9         Set R = Range(LeftCell, RightCell)
10        If Show Then
11            R.EntireColumn.Hidden = False
12            FormatCreditUsageSheet False
13            b.Characters.Text = WideCaption
14        Else
15            R.EntireColumn.Hidden = True
16            b.Characters.Text = NarrowCaption
17        End If

18        Left = LeftCell.offset(0, -1).Left
19        Width = RightCell.offset(, 1).Left - Left
20        If Width < 105 Then Width = 105
21        Top = shCreditUsage.Buttons("butMenu").Top
22        Height = shCreditUsage.Buttons("butMenu").Height

23        If b.Left <> Left Then
24            b.Left = Left
25        End If
26        If b.Width <> Width Then
27            b.Width = Width
28        End If
29        If b.Top <> Top Then
30            b.Top = Top
31        End If
32        If b.Height <> Height Then
33            b.Height = Height
34        End If

35        Exit Sub
ErrHandler:
36        SomethingWentWrong "#ShowHidePFEData (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetExtraTradesHeader
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : Sets the value of the header cell above the range ExtraTradeAmounts on the CreditUsage sheet.
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub SetExtraTradesHeader()
          Dim HeaderIs As String
          Dim HeaderShouldBe As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Select Case RangeFromSheet(shCreditUsage, "ExtraTradesAre").Value
              Case "IRS Airbus pays fixed EUR", "IRS Airbus receives fixed EUR"
3                 HeaderShouldBe = "Amounts EUR"
4             Case Else
5                 HeaderShouldBe = "Amounts USD"
6         End Select

7         With RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Cells(0, 1)
8             HeaderIs = .Value
9             If HeaderIs <> HeaderShouldBe Then
10                Set SPH = CreateSheetProtectionHandler(shCreditUsage)
11                .Value = HeaderShouldBe
12            End If
13        End With

14        Exit Sub
ErrHandler:
15        Throw "#SetExtraTradesHeader (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

