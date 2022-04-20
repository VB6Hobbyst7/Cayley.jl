Attribute VB_Name = "modNotionalBased"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modNotionalBased
' Author    : Philip Swannell
' Date      : 14-Jul-2015
'             Rewrite of code to implement our emulation of Banks' Credit line usage calculations
'             for those banks which use a "PV + % of Notional" type rules.
'             Code copes with Fx Options as well as Forwards.
'             TODO: Incorporate Rates products! (PGS 23 Sep 16.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NotionalBasedByFilters
' Author    : Philip Swannell
' Date      : 14-Jul-2015
' Purpose   : Wrapper to NotionalBasedFromTrades that gets trades from the Exteral Trades workbook,
'             appends any "Extra Trades" appends trades from the FutureTrades sheet, ages the trades then
'             calls the underlying method.
' PGS 12/10/16 Enhanced so that the (t=0) trade values are calculated using HW if ModelType is Multi-Currency Hull-White
'              Would be better to move Notional-based calculations to R... (now Julia)
' -----------------------------------------------------------------------------------------------------------------------
Function NotionalBasedByFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, ByVal BaseCCY As String, _
          ByVal FxNotionalPercentages, RatesNotionalPercentages, IncludeExtraTrades As Boolean, _
          IncludeAssetClasses As String, ExtraTradeLabels, ExtraTradeAmounts, IncludeFutureTrades As Boolean, _
          PortfolioAgeing As Double, TradesScaleFactor As Double, CurrenciesToInclude As String, ModelName As String, _
          TC As TradeCount, TimeEnd As Double, ByVal TimeGap As Double, ProductCreditLimits As String, _
          twb As Workbook, fts As Worksheet, ModelBareBones As Dictionary, ExtraTradesAre As String)

1         On Error GoTo ErrHandler

          Dim AnchorDate As Date
          Dim ExtraTradePVs As Variant
          Dim ExtraTradesJuliaFormat As Variant
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim Numeraire As String
          Dim TradePVs
          Dim TradesJuliaFormat

2         SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, IncludeExtraTrades, IncludeFxTrades, IncludeRatesTrades, False

3         AnchorDate = DictGet(ModelBareBones, "AnchorDate")
4         Numeraire = GetItem(ModelBareBones, "Numeraire")
5         TradesJuliaFormat = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, True, Numeraire, IncludeFxTrades, _
              IncludeRatesTrades, TradesScaleFactor, CurrenciesToInclude, False, TC, twb, fts, AnchorDate)
6         If sNRows(TradesJuliaFormat) > 1 Then
7             TradePVs = PortfolioValueHW(TradesJuliaFormat, ModelName, BaseCCY, True)
8         Else
9             TradePVs = CreateMissing()
10        End If

11        If IncludeExtraTrades Then
12            ExtraTradesJuliaFormat = ConstructExtraTrades(ExtraTradesAre, ModelBareBones, _
                  ExtraTradeAmounts, ExtraTradeLabels, False)
              'TODO should not the PV of the extra trades always be zero??? PGS 17 Jan 2022
13            ExtraTradePVs = PortfolioValueHW(ExtraTradesJuliaFormat, ModelName, BaseCCY, True)
14            TradePVs = Concatenate1DArrays(TradePVs, ExtraTradePVs)
15        End If

16        Force2DArrayRMulti ExtraTradeLabels, ExtraTradeAmounts, FxNotionalPercentages

17        NotionalBasedByFilters = NotionalBasedFromTrades(BaseCCY, FxNotionalPercentages, RatesNotionalPercentages, _
              ModelBareBones, TradesJuliaFormat, ExtraTradesJuliaFormat, TimeEnd, TimeGap, TradePVs, "3Cols")

18        Exit Function
ErrHandler:
19        NotionalBasedByFilters = "#NotionalBasedByFilters (line " & CStr(Erl) & "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NotionalBasedFromTrades
' Author    : Philip Swannell
' Date      : 14-Jul-2015, updated Nov 2016 to handle Rates trades.
' Purpose   : FxNotionalPercentages to be a table giving line utilisation as a percentage of notional (data for
' Bank of Tokyo)
'  1  .12
'  2  .17
'  3  .21
'  4  .24
'  5  .27
'  7  .32
' Line utilisation is then:
' Trade PV + Sum over trades(percentage of preferred currency notional * FxSpot rate to BaseCcy)
' There is no netting between trades.  RatesNotionalPercentages are passed in as per example below
''Tenor'  'USD'    'EUR'    'Other'
'8.3333  0.00071  0.00138  0.00138
'0.1666  0.00122  0.00277  0.00277
'0.25    0.00214  0.00415  0.00415
'0.5     0.00429  0.0083   0.0083
'0.75    0.00643  0.0125   0.0125
'1       0.00858  0.0166   0.0166
'2       0.0221   0.0171   0.0171
'3       0.0364   0.0309   0.0309
'4       0.0591   0.0443   0.0443
'5       0.0778   0.0576   0.0576
'
'WhatToReturn can take values NotionalsForCap, TotalNotionalForCap, 3Cols, LineUse

' -----------------------------------------------------------------------------------------------------------------------
Function NotionalBasedFromTrades(ByVal BaseCCY As String, FxNotionalPercentages, RatesNotionalPercentages, _
          ModelBareBones As Dictionary, TradesJuliaFormat, ExtraTradesJuliaFormat, _
          TimeEnd As Double, ByVal TimeGap As Double, ByVal TradePVs As Variant, WhatToReturn)

1         On Error GoTo ErrHandler

          Const NPInterpMethod = "Linear"
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NumTrades As Long
          Dim Widths, Heights        'Thinking of each trade as a rectangle, height = line usage, width = maturity
          Dim AnchorDate
          Dim Ccy
          Dim cn_Currency
          Dim cn_Dates
          Dim cn_EndDate
          Dim cn_Notional
          Dim cn_PayAmortNotionals
          Dim cn_PayCurrency
          Dim cn_PayNotional
          Dim cn_PayNotionals
          Dim cn_ReceiveAmortNotionals
          Dim cn_ReceiveCurrency
          Dim cn_ReceiveNotional
          Dim cn_ReceiveNotionals
          Dim cn_ValuationFunction
          Dim EndDate
          Dim FxNPLeftCol
          Dim FxNPRightCol
          Dim Headers As Variant
          Dim Notional
          Dim NotionalsForCap As Variant
          Dim PayCurrency
          Dim PayNotional
          Dim RatesNPHeaders
          Dim RatesNPLeftCol
          Dim ReceiveCurrency
          Dim ReceiveNotional
          Dim ThisBaseCCyNotional
          Dim ThisNotionalPercentage
          Dim ThisPV As Double
          Dim TradesArray
          Dim Zeros As Variant

          Dim ValuationFunction As String
          'Validate inputs e.g. the 2 arrays of notional percentages

          'NotionalBased is fast to run, always use at least 400 observations...
2         If TimeGap > TimeEnd / 400 Then TimeGap = TimeEnd / 400

3         AnchorDate = DictGet(ModelBareBones, "AnchorDate")
4         FxNPLeftCol = sSubArray(FxNotionalPercentages, 1, 1, , 1)
5         FxNPRightCol = sSubArray(FxNotionalPercentages, 1, 2, , 1)
6         RatesNPHeaders = sArrayTranspose(sSubArray(RatesNotionalPercentages, 1, 1, 1))
7         RatesNPLeftCol = sSubArray(RatesNotionalPercentages, 2, 1, , 1)

8         If sNRows(TradesJuliaFormat) > 1 Then
9             NumTrades = sNRows(TradesJuliaFormat) - 1
10        Else
11            NumTrades = 0
12        End If
13        If sNRows(ExtraTradesJuliaFormat) > 1 Then
14            NumTrades = NumTrades + sNRows(ExtraTradesJuliaFormat) - 1
15        End If

16        Zeros = sReshape(0, NumTrades, 1)
17        Widths = Zeros
18        Heights = Zeros
19        NotionalsForCap = Zeros

20        k = 0
21        For j = 1 To 2
22            TradesArray = Choose(j, TradesJuliaFormat, ExtraTradesJuliaFormat)
23            If sNRows(TradesArray) > 1 Then
24                Headers = sArrayTranspose(sSubArray(TradesArray, 1, 1, 1))
25                cn_ValuationFunction = sMatch("ValuationFunction", Headers)
26                If Not IsNumber(cn_ValuationFunction) Then
27                    Throw "Cannot find ValuationFunction in header row of trades"
28                End If
29                cn_ReceiveCurrency = sMatch("ReceiveCurrency", Headers)
30                If Not IsNumber(cn_ReceiveCurrency) Then
31                    Throw "Cannot find ReceiveCurrency in header row of trades"
32                End If
33                cn_PayCurrency = sMatch("PayCurrency", Headers)
34                If Not IsNumber(cn_PayCurrency) Then
35                    Throw "Cannot find PayCurrency in header row of trades"
36                End If
37                cn_ReceiveNotional = sMatch("ReceiveNotional", Headers)
38                If Not IsNumber(cn_ReceiveNotional) Then
39                    Throw "Cannot find ReceiveNotional in header row of trades"
40                End If
41                cn_PayNotional = sMatch("PayNotional", Headers)
42                If Not IsNumber(cn_PayNotional) Then
43                    Throw "Cannot find PayNotional in header row of trades"
44                End If
45                cn_EndDate = sMatch("EndDate", Headers)
46                If Not IsNumber(cn_EndDate) Then
47                    Throw "Cannot find EndDate in header row of trades"
48                End If

49                cn_Currency = sMatch("Currency", Headers)
50                If Not IsNumber(cn_Currency) Then
51                    If j = 1 Then
52                        Throw "Cannot find Currency in header row of trades"
53                    Else
54                        cn_Currency = 0
55                    End If
56                End If
57                cn_Notional = sMatch("Notional", Headers)
58                If Not IsNumber(cn_Notional) Then
59                    If j = 1 Then
60                        Throw "Cannot find Notional in header row of trades"
61                    Else
62                        cn_Notional = 0
63                    End If
64                End If
65                cn_ReceiveAmortNotionals = sMatch("ReceiveAmortNotionals", Headers)
66                If Not IsNumber(cn_ReceiveAmortNotionals) Then
67                    If j = 1 Then
68                        Throw "Cannot find ReceiveAmortNotionals in header row of trades"
69                    Else
70                        cn_ReceiveAmortNotionals = 0
71                    End If
72                End If
73                cn_PayAmortNotionals = sMatch("PayAmortNotionals", Headers)
74                If Not IsNumber(cn_PayAmortNotionals) Then
75                    If j = 1 Then
76                        Throw "Cannot find PayAmortNotionals in header row of trades"
77                    Else
78                        cn_PayAmortNotionals = 0
79                    End If
80                End If
81                cn_Dates = sMatch("Dates", Headers)
82                If Not IsNumber(cn_Dates) Then
83                    If j = 1 Then
84                        Throw "Cannot find Dates in header row of trades"
85                    Else
86                        cn_Dates = 0
87                    End If
88                End If
89                cn_ReceiveNotionals = sMatch("ReceiveNotionals", Headers)
90                If Not IsNumber(cn_ReceiveNotionals) Then
91                    If j = 1 Then
92                        Throw "Cannot find ReceiveNotionals in header row of trades"
93                    Else
94                        cn_ReceiveNotionals = 0
95                    End If
96                End If
97                cn_PayNotionals = sMatch("PayNotionals", Headers)
98                If Not IsNumber(cn_PayNotionals) Then
99                    If j = 1 Then
100                       Throw "Cannot find PayNotionals in header row of trades"
101                   Else
102                       cn_PayNotionals = 0
103                   End If
104               End If

105           End If

106           For i = 1 To sNRows(TradesArray) - 1
107               k = k + 1
108               EndDate = TradesArray(i + 1, cn_EndDate)
109               ValuationFunction = TradesArray(i + 1, cn_ValuationFunction)
110               ReceiveCurrency = TradesArray(i + 1, cn_ReceiveCurrency)
111               PayCurrency = TradesArray(i + 1, cn_PayCurrency)
112               ReceiveNotional = TradesArray(i + 1, cn_ReceiveNotional)
113               PayNotional = TradesArray(i + 1, cn_PayNotional)
114               ThisPV = TradePVs(k)

115               Widths(k, 1) = (EndDate - AnchorDate) / 365

116               Select Case ValuationFunction

                      Case "FxForward", "CrossCurrencySwap", "FxForwardStrip"

117                       If ValuationFunction = "FxForwardStrip" Then
                              'We assume that this is a booking of an FxSwap, so there should be two elements encoded in the trades Dates, ReceiveNotionals and PayNotionals attributes
118                           ParseFxForwardStripAsFxSwap CStr(TradesArray(i + 1, cn_Dates)), _
                                  CStr(TradesArray(i + 1, cn_ReceiveNotionals)), _
                                  CStr(TradesArray(i + 1, cn_PayNotionals)), _
                                  EndDate, ReceiveNotional, PayNotional
119                       End If

120                       If PayCurrency = BaseCCY Then
121                           ThisBaseCCyNotional = PayNotional
122                       ElseIf ReceiveCurrency = BaseCCY Then
123                           ThisBaseCCyNotional = ReceiveNotional
124                       Else
                              Dim PrefCcy As String
125                           PrefCcy = PreferredCcy(ReceiveCurrency, PayCurrency, BaseCCY)
126                           If PrefCcy = PayCurrency Then
127                               ThisBaseCCyNotional = PayNotional * _
                                      MyFxPerBaseCcy(PrefCcy, BaseCCY, ModelBareBones)
128                           ElseIf PrefCcy = ReceiveCurrency Then
129                               ThisBaseCCyNotional = ReceiveNotional * _
                                      MyFxPerBaseCcy(PrefCcy, BaseCCY, ModelBareBones)
130                           Else
131                               Throw "Assertion failed."
132                           End If
133                       End If
134                       ThisNotionalPercentage = FirstElementOf(sInterp(FxNPLeftCol, FxNPRightCol, _
                              Widths(k, 1), NPInterpMethod, "FF"))
135                       NotionalsForCap(k, 1) = Abs(ThisBaseCCyNotional)
136                       Heights(k, 1) = ThisPV + ThisNotionalPercentage * Abs(ThisBaseCCyNotional)
137                       If Heights(k, 1) < 0 Then Heights(k, 1) = 0
138                   Case "FxOption"
139                       Ccy = TradesArray(i + 1, cn_Currency)
140                       Notional = TradesArray(i + 1, cn_Notional)

141                       ThisBaseCCyNotional = Notional * MyFxPerBaseCcy(Ccy, BaseCCY, ModelBareBones)
                          ' Short option positions do count for Notional Cap.
142                       NotionalsForCap(k, 1) = Abs(ThisBaseCCyNotional)
143                       If Notional < 0 Then        'Bank is short the option so no credit risk
144                           ThisBaseCCyNotional = 0
145                       End If
146                       ThisNotionalPercentage = FirstElementOf(sInterp(FxNPLeftCol, FxNPRightCol, _
                              Widths(k, 1), NPInterpMethod, "FF"))
147                       Heights(k, 1) = ThisPV + ThisNotionalPercentage * ThisBaseCCyNotional
148                       If Heights(k, 1) < 0 Then Heights(k, 1) = 0
149                   Case "InterestRateSwap"
                          'Cope with amortising trades by (usually conservative) choice of first notional
150                       Ccy = TradesArray(i + 1, cn_ReceiveCurrency)
151                       Notional = TradesArray(i + 1, cn_ReceiveNotional)
152                       If cn_ReceiveAmortNotionals <> 0 Then
153                           If VarType(TradesArray(i + 1, cn_ReceiveAmortNotionals)) = vbString Then
154                               Notional = CDbl(sStringBetweenStrings(TradesArray(i + 1, cn_ReceiveAmortNotionals), , ";"))
155                           End If
156                       End If

157                       If Ccy = BaseCCY Then
158                           ThisBaseCCyNotional = Abs(Notional)
159                       Else
160                           ThisBaseCCyNotional = Abs(Notional) * MyFxPerBaseCcy(Ccy, BaseCCY, ModelBareBones)
161                       End If
162                       NotionalsForCap(k, 1) = ThisBaseCCyNotional
                          Dim MatchID
                          Dim yArray
163                       MatchID = sMatch(Ccy, RatesNPHeaders)
164                       If Not IsNumber(MatchID) Then
165                           MatchID = sMatch("Other", RatesNPHeaders)
166                       End If
167                       yArray = sSubArray(RatesNotionalPercentages, 2, MatchID, , 1)
168                       ThisNotionalPercentage = FirstElementOf(sInterp(RatesNPLeftCol, yArray, _
                              Widths(k, 1), NPInterpMethod, "FF"))
169                       Heights(k, 1) = ThisPV + ThisNotionalPercentage * ThisBaseCCyNotional
170                       If Heights(k, 1) < 0 Then Heights(k, 1) = 0
171                   Case Else
172                       Throw "ValuationFunction (" & ValuationFunction & ") not handled"
173               End Select

174           Next i
175       Next j

176       If WhatToReturn = "NotionalsForCap" Then
177           NotionalBasedFromTrades = NotionalsForCap
178           Exit Function
179       ElseIf WhatToReturn = "TotalNotionalForCap" Then
180           NotionalBasedFromTrades = sColumnSum(NotionalsForCap)(1, 1)
181           Exit Function
182       End If

          Dim LineUseAtObservationDates
          Dim LookupArray
          Dim MaxDate
          Dim ObservationDates
          Dim ObservationTimes
183       ObservationTimes = sGrid(0, TimeEnd, 1 + TimeEnd / TimeGap)
184       ObservationDates = sArrayAdd(AnchorDate, sArrayMultiply(ObservationTimes, 365))

185       If NumTrades > 0 Then
186           LookupArray = sArrayRange(Widths, Heights)
187           LookupArray = sSortMerge(LookupArray, 1, 2, "Sum")
188           MaxDate = sColumnMax(Widths)(1, 1)
189           LookupArray = sArrayStack(LookupArray, sArrayRange(MaxDate + 0.1, 0))
190           For i = sNRows(LookupArray) - 1 To 1 Step -1
191               LookupArray(i, 2) = LookupArray(i, 2) + LookupArray(i + 1, 2)
192           Next i
193           LineUseAtObservationDates = ThrowIfError(sInterp(sSubArray(LookupArray, 1, 1, , 1), _
                  sSubArray(LookupArray, 1, 2, , 1), ObservationTimes, "FlatToRight", "FF"))
194       Else
195           LineUseAtObservationDates = sReshape(0, sNRows(ObservationDates), 1)
196       End If

197       If WhatToReturn = "3Cols" Then
198           NotionalBasedFromTrades = sArrayRange(ObservationDates, ObservationTimes, LineUseAtObservationDates)
199       ElseIf WhatToReturn = "LineUse" Then
200           NotionalBasedFromTrades = LineUseAtObservationDates
201       Else
202           Throw "Unrecognised value for WhatToReturn"
203       End If

204       Exit Function
ErrHandler:
205       Throw "#NotionalBasedFromTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseFxForwardStripAsFxSwap
' Author     : Philip Swannell
' Date       : 08-Mar-2022
' Purpose    : Sub of NotionalBasedFromTrades, and used to treat FxForwardStrip with two tradelets as an FxSwap, for which
'              we treat as having the same line usage as a forward to the far date.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseFxForwardStripAsFxSwap(Dates As String, ReceiveNotionals As String, PayNotionals As String, ByRef EndDate, ByRef ReceiveNotional, ByRef PayNotional)

          Dim NumForwards As Long
1         On Error GoTo ErrHandler
2         NumForwards = Len(Dates) - Len(Replace(Dates, ";", "")) + 1
3         If NumForwards <> 2 Then
4             Throw "FxForwardStrip must contain 2 FxForward trades (i.e. represent an Fx Swap) but trade contains " + CStr(NumForwards) + " FxForward trades"
5         End If
6         EndDate = CDate(Mid(Dates, InStrRev(Dates, ";") + 1))
7         ReceiveNotional = CDbl(Mid(ReceiveNotionals, InStrRev(ReceiveNotionals, ";") + 1))
8         PayNotional = CDbl(Mid(PayNotionals, InStrRev(PayNotionals, ";") + 1))
9         Exit Sub
ErrHandler:
10        Throw "#ParseFxForwardStripAsFxSwap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NBSolverFromFilters
' Author    : Philip Swannell
' Date      : 18-Oct-2016
' Purpose   : Hand-crafted solver to solve for trade headroom in the case of banks using
'             Notional-based methodology. Much faster than using the Naive methods
'             TradeHeadroomSolverLnFx or Solve1to5Naive which treat the PFE sheet as a "black box"
'             and change inputs until a particular output reaches a target level.
'
'             Most inputs to the function are as used by many other methods. UnitHedgeAmounts
'             is an array with as many rows as the column array HedgeLabels. UnitHedgeAmounts
'             may have multiple columns in which case the ith element of the return is the "multiplier"
'             such that Portfolio(ExistingTrades + multiplier x ith column of UnitHedgeAmounts)
'             exhauts the credit lines.
' -----------------------------------------------------------------------------------------------------------------------
Function NBSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
          IncludeAssetClasses As String, HedgeLabels, UnitHedgeAmounts, PortfolioAgeing As Double, _
          FlipTrades As Boolean, ModelName As String, ByVal TimeGap As Double, TimeEnd As Double, _
          FxNotionalPercentages As Variant, RatesNotionalPercentages As Variant, LimitTimes, CreditLimits, _
          CreditInterpMethod, BaseCCY As String, TradesScaleFactor As Double, ByRef ProfileWithET, _
          ByRef ProfileWithoutET, CurrenciesToInclude As String, TC As TradeCount, ProductCreditLimits As String, _
          DoNotionalCap As Boolean, ByRef NotionalCapApplies As Boolean, NotionalCapForNewTrades As Double, _
          twb As Workbook, fts As Worksheet, ModelBareBones As Dictionary, ExtraTradesAre As String)

          Dim AnchorDate As Date
          Dim ExistingTradesProfile
          Dim ExtraTradesJuliaFormat As Variant
          Dim HedgeTradePVs
          Dim i As Long
          Dim IncludeExtraTrades As Boolean
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim InterpolatedLines
          Dim Multipliers
          Dim NumTrades As Long
          Dim ObservationDates
          Dim ObservationTimes
          Dim TheseHedgeAmounts
          Dim ThisMultiplier
          Dim TradesJuliaFormat
          Dim UnitHedgeTradesProfile
          Const NumSims = 1        ' Does not matter
          Const PFEPercentile = 0.95        'Does not matter
          Const Methodology = "Notional Based"
          Const ShortfallOrQuantile = "-"        'Does not matter

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti HedgeLabels, UnitHedgeAmounts, FxNotionalPercentages
3         SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, True, IncludeFxTrades, IncludeRatesTrades, False

4         AnchorDate = DictGet(ModelBareBones, "AnchorDate")
5         TradesJuliaFormat = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, FlipTrades, NumeraireFromMDWB, IncludeFxTrades, _
              IncludeRatesTrades, TradesScaleFactor, CurrenciesToInclude, False, TC, twb, fts, AnchorDate)
6         NumTrades = TC.NumIncluded
7         If NumTrades <> sNRows(TradesJuliaFormat) - 1 Then
8             Throw "Assertion Failed - two methods of counting trades differ"
9         End If
          'NotionalBased is fast to run, always use at least 400 observations...
10        If TimeGap > TimeEnd / 400 Then TimeGap = TimeEnd / 400

11        ObservationTimes = sGrid(0, TimeEnd, 1 + TimeEnd / TimeGap)
12        ObservationDates = sArrayAdd(AnchorDate, sArrayMultiply(ObservationTimes, 365))

13        If NumTrades = 0 Then
14            ExistingTradesProfile = sReshape(0, sNRows(ObservationTimes), 1)
15        Else
16            ExistingTradesProfile = ThrowIfError(PFEProfileFromFiltersPCL(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
                  IncludeFutureTrades, IncludeAssetClasses, IncludeExtraTrades, HedgeLabels, UnitHedgeAmounts, _
                  PortfolioAgeing, FlipTrades, TradesScaleFactor, BaseCCY, ModelName, NumSims, TimeGap, TimeEnd, _
                  Methodology, PFEPercentile, ShortfallOrQuantile, FxNotionalPercentages, RatesNotionalPercentages, _
                  CurrenciesToInclude, ModelName, TC, ProductCreditLimits, twb, True, ModelBareBones, ExtraTradesAre))

17        End If

18        ProfileWithoutET = ExistingTradesProfile

          'No need to value them, they have zero PV by construction
19        HedgeTradePVs = Repeat(0, sNRows(UnitHedgeAmounts))
20        InterpolatedLines = ThrowIfError(sInterp(LimitTimes, CreditLimits, ObservationTimes, CreditInterpMethod, "FF"))

21        Multipliers = sReshape(0, sNCols(UnitHedgeAmounts), 1)

          Dim MaxMultiplier

22        For i = 1 To sNCols(UnitHedgeAmounts)
23            TheseHedgeAmounts = sSubArray(UnitHedgeAmounts, 1, i, , 1)
24            ExtraTradesJuliaFormat = ConstructExtraTrades(ExtraTradesAre, ModelBareBones, _
                  TheseHedgeAmounts, HedgeLabels, False)
25            UnitHedgeTradesProfile = NotionalBasedFromTrades(BaseCCY, FxNotionalPercentages, _
                  RatesNotionalPercentages, ModelBareBones, Empty, ExtraTradesJuliaFormat, TimeEnd, _
                  TimeGap, HedgeTradePVs, "LineUse")
26            ThisMultiplier = sArrayDivide(sArraySubtract(InterpolatedLines, _
                  sSubArray(ExistingTradesProfile, 1, 3, , 1)), _
                  UnitHedgeTradesProfile)
                  
              'UnitHedgeTradesProfile is zero beyond time = longest hedge trade, so the elements of ThisMultiplier _
               are error strings beyond that point, so filtered out by the call to sMinOfNums. _
               In this way we pay attention only to the credit limits up to the maximum maturity of the Unit Hedges _
               which is what we want to do.
                  
27            ThisMultiplier = sMinOfNums(ThisMultiplier)

28            If ThisMultiplier < 0 Then ThisMultiplier = 0        'lines are full
29            If DoNotionalCap Then
30                MaxMultiplier = NotionalCapForNewTrades / sColumnSum(sSubArray(UnitHedgeAmounts, 1, i, , 1))(1, 1)
31                If ThisMultiplier > MaxMultiplier Then
32                    NotionalCapApplies = True
33                    Multipliers(i, 1) = MaxMultiplier
34                Else
35                    Multipliers(i, 1) = ThisMultiplier
36                End If
37            Else
38                Multipliers(i, 1) = ThisMultiplier
39            End If
40        Next i

          'Show the Profile for the "last" of the possibilities for unitHedgeTrades
41        ProfileWithET = sArrayRange(ObservationDates, ObservationTimes, _
              sArrayAdd(sSubArray(ExistingTradesProfile, 1, 3, , 1), _
              sArrayMultiply(sTake(Multipliers, -1), UnitHedgeTradesProfile)))
42        NBSolverFromFilters = Multipliers
43        Exit Function
ErrHandler:
44        NBSolverFromFilters = "#NBSolverFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnPackNotionalPercentages
' Author    : Philip Swannell
' Date      : 16-Nov-2016
' Purpose   : Unpacks notional weights, which are stored as array strings, converts left
'             column from tenor to time and validates the data.
' -----------------------------------------------------------------------------------------------------------------------
Function UnPackNotionalPercentages(ArrayString As String, IsRates As Boolean)
          Dim AllCcys As Variant
          Dim ArrayToCheck As Variant
          Dim ArrayToCheckOrig As Variant
          Dim i As Long
          Dim j As Long
          Dim LeftCol As Variant
          Dim NameForError As String
          Dim Percentages As Variant
          Dim Times As Variant

1         On Error GoTo ErrHandler
2         NameForError = IIf(IsRates, "Rates Notional Weights", "Rates Notional Weights")

3         If Left(ArrayString, 1) <> "{" Or Right(ArrayString, 1) <> "}" Then
4             Throw NameForError & " must be saved in the lines workbook as a " & _
                  "string starting with ""{"" and ending with ""}"""
5         End If
6         ArrayToCheck = sParseArrayString(ArrayString)
7         ArrayToCheckOrig = ArrayToCheck
8         If VarType(ArrayToCheck) = vbString Then
9             Throw NameForError & " appears in the lines workbook in invalid format. Use the ""Double-click "" " & _
                  "to edit functionality of the lines workbook to ensure that the data is saved in a valid format"
10        End If

11        If IsRates Then
12            If sNRows(ArrayToCheck) < 2 Then
13                Throw NameForError & " must have at least two rows"
14            ElseIf sNCols(ArrayToCheck) < 2 Then
15                Throw NameForError & " must have at least two columns"
16            End If
17            Percentages = sSubArray(ArrayToCheck, 2, 2)
18        Else
19            If sNRows(ArrayToCheck) < 2 Then
20                Throw NameForError & " must have at least two rows"
21            ElseIf sNCols(ArrayToCheck) <> 2 Then
22                Throw NameForError & " must have two columns"
23            End If
24            Percentages = sSubArray(ArrayToCheck, 1, 2)
25        End If

26        LeftCol = sSubArray(ArrayToCheck, IIf(IsRates, 2, 1), 1, , 1)

27        Times = 0
28        On Error Resume Next
29        Times = TenorToTime(LeftCol)
30        On Error GoTo ErrHandler
31        If sArraysIdentical(Times, 0) Then
32            Throw "The tenors in " & NameForError & " must end in either 'M' or 'Y' eg '6M' or '5Y'"
33        End If

34        For i = 2 To sNRows(Times)
35            If Times(i, 1) <= Times(i - 1, 1) Then
36                Throw "The tenors in " & NameForError & " are not correctly ordered"
37            End If
38        Next i

39        For j = 1 To sNCols(Percentages)
40            For i = 1 To sNRows(Percentages)
41                If Not IsNumber(Percentages(i, j)) Then
42                    Throw "The weights in " & NameForError & " must be numbers"
43                ElseIf Percentages(i, j) < 0 Then
44                    Throw "The weights in " & NameForError & " must be non-negative numbers"
45                ElseIf i > 1 Then
46                    If Percentages(i, j) < Percentages(i - 1, j) Then
47                        Throw "The weights in " & NameForError & " must increase with time"
48                    End If
49                End If
50            Next i
51        Next j

52        If IsRates Then
53            If CStr(ArrayToCheck(1, 1)) <> "Tenor" Then
54                Throw "Top-left element of " & NameForError & " must be 'Tenor'"
55            End If
56            For j = 2 To sNCols(ArrayToCheck)
57                AllCcys = sCurrencies(False, False)
58                If ArrayToCheck(1, j) <> "Other" Then
59                    If Not IsNumber(sMatch(ArrayToCheck(1, j), AllCcys)) Then
60                        Throw "items in the top row of " & NameForError & " must be valid currency codes or 'Other'"
61                    End If
62                End If
63            Next j
64        End If

65        For i = 1 To sNRows(Times)
66            ArrayToCheck(i + IIf(IsRates, 1, 0), 1) = Times(i, 1)
67        Next i

68        UnPackNotionalPercentages = ArrayToCheck

69        Exit Function
ErrHandler:
70        Throw "#UnPackNotionalPercentages (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

