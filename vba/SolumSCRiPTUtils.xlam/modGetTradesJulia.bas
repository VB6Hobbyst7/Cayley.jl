Attribute VB_Name = "modGetTradesJulia"
Option Explicit

Sub Test_GTIJF()

          Const FilterBy1 = "Trade Id"
          Const Filter1Value = 1170525
          Const FilterBy2 = "None"
          Const Filter2Value = "None"
          Const IncludeFutureTrades As Boolean = False
          Const PortfolioAgeing As Double = 0
          Const FlipTrades As Boolean = False
          Const Numeraire = "EUR"
          Const WithFxTrades = True
          Const WithRatesTrades = True
          Const TradesScaleFactor = 1
          Const CurrenciesToInclude As String = "All"
          Const Compress As Boolean = False
          Dim TC As TradeCount
          Dim twb As Workbook
          Dim fts As Worksheet
          Dim AnchorDate As Date
          Dim Res
          
1         On Error GoTo ErrHandler
2         Set twb = Application.Workbooks("CayleyTrades.xlsx")
3         AnchorDate = Date

4         Res = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, _
              PortfolioAgeing, FlipTrades, Numeraire, _
              WithFxTrades, WithRatesTrades, TradesScaleFactor, _
              CurrenciesToInclude, Compress, _
              TC, twb, fts, AnchorDate)

5         g sArrayTranspose(Res), ExMthdSpreadsheet

6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#Test_GTIJF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetTradesInJuliaFormat
' Author    : Philip Swannell
' Date      : 03-Mar-2022
' Purpose   : Takes trades from the trades workbook and puts then in the format needed
'             for Julia. Also applies PortfolioAgeing by amending the trades' maturity dates.
'             Differences between format required by the R code and the XVA code are described
'             at https://github.com/SolumXplain/XVA/wiki/Trade-files.
'This version is for use with Trade workbooks as they are for the 2022 Cayley project. The workbooks contain the contents
'of .csv files, though the headers are morphed by code in the Cayley workbook (see worksheet StaticData in the Cayley workbook)
' -----------------------------------------------------------------------------------------------------------------------
Function GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
          PortfolioAgeing As Double, FlipTrades As Boolean, Numeraire As String, _
          WithFxTrades As Boolean, WithRatesTrades As Boolean, TradesScaleFactor As Double, _
          CurrenciesToInclude As String, Compress As Boolean, _
          TC As TradeCount, twb As Workbook, fts As Worksheet, AnchorDate As Date)

1         On Error GoTo ErrHandler
          'Frequently call with same arguments as last time, so use cacheing. _
           Note that NumTrades should not be part of the CheckSum.
          'Cacheing is valid as long as the trades in the trades workbook and the _
           FutureTrades sheet have not changed since the previous call. So we have _
           method FlushStatics which gets called after that might have happened.
          Dim ThisCheckSum As String
          Static PreviousCheckSum As String
          Static PreviousResult As Variant
          Static PreviousNumTrades As Long
          Static PreviousTC As TradeCount

2         ThisCheckSum = CStr(FilterBy1) & "," & CStr(Filter1Value) & "," & CStr(FilterBy2) & "," & _
              CStr(Filter2Value) & "," & CStr(IncludeFutureTrades) & "," & CStr(PortfolioAgeing) & "," & _
              CStr(FlipTrades) & "," & CStr(Numeraire) & "," & CStr(WithFxTrades) & "," & _
              CStr(WithRatesTrades) & "," & CStr(TradesScaleFactor) & "," & CStr(CurrenciesToInclude) & "," & _
              CStr(Compress) & "," & twb.FullName
              
3         If ThisCheckSum = PreviousCheckSum Then
4             If Not (IsEmpty(PreviousResult)) Then
5                 GetTradesInJuliaFormat = PreviousResult
6                 TC = PreviousTC
7                 Exit Function
8             End If
9         End If

          Dim AnySchedules As Boolean
          Dim Ccy As String
          Dim ChooseVector As Variant
          Dim ErrString As String
          Dim i As Long
          Dim InTrades As Variant
          Dim IsCall As Boolean
          Dim IsCallOnPrimCur As Boolean
          Dim k As Long
          Dim LongTheOpt As Boolean
          Dim Notional As Double
          Dim NumOutTrades As Long
          Dim OutTrades() As Variant
          Dim Strike As Double
          Dim ThisScaleFactor As Double
          Dim TradeID As String
          Dim VF As String
          Dim HasSchedule As Boolean

          Const nOut_TradeID = 1
          Const nOut_ValuationFunction = 2
          Const nOut_Counterparty = 3
          Const nOut_StartDate = 4
          Const nOut_EndDate = 5
          Const nOut_Currency = 6
          Const nOut_ReceiveCurrency = 7
          Const nOut_PayCurrency = 8
          Const nOut_Notional = 9
          Const nOut_ReceiveNotional = 10
          Const nOut_PayNotional = 11
          Const nOut_Strike = 12
          Const nOut_isCall = 13
          Const nOut_PayAmortNotionals = 14
          Const nOut_PayBDC = 15
          Const nOut_PayCoupon = 16
          Const nOut_PayDCT = 17
          Const nOut_PayFrequency = 18
          Const nOut_PayIndex = 19
          Const nOut_ReceiveAmortNotionals = 20
          Const nOut_ReceiveBDC = 21
          Const nOut_ReceiveCoupon = 22
          Const nOut_ReceiveDCT = 23
          Const nOut_ReceiveFrequency = 24
          Const nOut_ReceiveIndex = 25
          Const nOut_Dates = 26
          Const nOut_ReceiveNotionals = 27
          Const nOut_PayNotionals = 28
          Const nOut_NumeraireEquals = 29

          'Changing the headers we look for? Then amend method CheckTradesWorkbook
          Const HeaderNames = "Trade Id,Product Type,Counterparty Parent,Settle Date,Maturity Date,Prim Cur,Prim Amt,Sec Cur,Sec Amt,FX Prim Amount,FX Sec Amount,FX Far Prim Amount,FX Far Sec Amount,Buy Sell,Prin Sched Type,Pay Principal,Pay Ccy,Pay Type,Pay Fixed Rate,Pay Floating Rate,Pay Freq,Pay Daycount,Pay BDC,Rcv Principal,Rcv Ccy,Rcv Type,Rcv Fixed Rate,Rcv Floating Rate,Rcv Freq,Rcv Daycount,Rcv BDC,Rate Index Spread,TradeIsFrom"
          
          'These variables must be in synch with the constant HeaderNames
          Const nIn_TradeId = 1
          Const nIn_ProductType = 2
          Const nIn_CounterpartyParent = 3
          Const nIn_SettleDate = 4
          Const nIn_MaturityDate = 5
          Const nIn_PrimCur = 6
          Const nIn_PrimAmt = 7
          Const nIn_SecCur = 8
          Const nIn_SecAmt = 9
          Const nIn_FXPrimAmount = 10
          Const nIn_FXSecAmount = 11
          Const nIn_FXFarPrimAmount = 12
          Const nIn_FXFarSecAmount = 13
          Const nIn_BuySell = 14
          Const nIn_PrinSchedType = 15
          Const nIn_PayPrincipal = 16
          Const nIn_PayCcy = 17
          Const nIn_PayType = 18
          Const nIn_PayFixedRate = 19
          Const nIn_PayFloatingRate = 20
          Const nIn_PayFreq = 21
          Const nIn_PayDaycount = 22
          Const nIn_PayBDC = 23
          Const nIn_RcvPrincipal = 24
          Const nIn_RcvCcy = 25
          Const nIn_RcvType = 26
          Const nIn_RcvFixedRate = 27
          Const nIn_RcvFloatingRate = 28
          Const nIn_RcvFreq = 29
          Const nIn_RcvDaycount = 30
          Const nIn_RcvBDC = 31
          Const nIn_RateIndexSpread = 32
          Const nIn_TradeIsFrom = 33
          
10        On Error GoTo ErrHandler

11        ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, PortfolioAgeing, False, WithFxTrades, WithRatesTrades, CurrenciesToInclude, TC, twb, fts, AnchorDate)
12        InTrades = GetColumnsFromTradesWorkbook(HeaderNames, IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)

13        If Compress Then
              'FXSwaps represented as two rows
14            For i = 1 To sNRows(ChooseVector)
15                If ChooseVector(i, 1) Then
16                    If InTrades(i, nIn_ProductType) = "FXSwap" Then
17                        NumOutTrades = NumOutTrades + 2
18                    Else
19                        NumOutTrades = NumOutTrades + 1
20                    End If
21                End If
22            Next i
23        Else
24            NumOutTrades = sArrayCount(ChooseVector)
25        End If

26        ReDim OutTrades(1 To NumOutTrades + 1, 1 To 29)
27        OutTrades(1, nOut_TradeID) = "TradeID"
28        OutTrades(1, nOut_ValuationFunction) = "ValuationFunction"
29        OutTrades(1, nOut_Counterparty) = "Counterparty"
30        OutTrades(1, nOut_StartDate) = "StartDate"
31        OutTrades(1, nOut_EndDate) = "EndDate"
32        OutTrades(1, nOut_Currency) = "Currency"
33        OutTrades(1, nOut_ReceiveCurrency) = "ReceiveCurrency"
34        OutTrades(1, nOut_PayCurrency) = "PayCurrency"
35        OutTrades(1, nOut_Notional) = "Notional"
36        OutTrades(1, nOut_ReceiveNotional) = "ReceiveNotional"
37        OutTrades(1, nOut_PayNotional) = "PayNotional"
38        OutTrades(1, nOut_Strike) = "Strike"
39        OutTrades(1, nOut_isCall) = "IsCall"
40        OutTrades(1, nOut_PayAmortNotionals) = "PayAmortNotionals"
41        OutTrades(1, nOut_PayBDC) = "PayBDC"
42        OutTrades(1, nOut_PayCoupon) = "PayCoupon"
43        OutTrades(1, nOut_PayDCT) = "PayDCT"
44        OutTrades(1, nOut_PayFrequency) = "PayFrequency"
45        OutTrades(1, nOut_PayIndex) = "PayIndex"
46        OutTrades(1, nOut_ReceiveAmortNotionals) = "ReceiveAmortNotionals"
47        OutTrades(1, nOut_ReceiveBDC) = "ReceiveBDC"
48        OutTrades(1, nOut_ReceiveCoupon) = "ReceiveCoupon"
49        OutTrades(1, nOut_ReceiveDCT) = "ReceiveDCT"
50        OutTrades(1, nOut_ReceiveFrequency) = "ReceiveFrequency"
51        OutTrades(1, nOut_ReceiveIndex) = "ReceiveIndex"
52        OutTrades(1, nOut_Dates) = "Dates"
53        OutTrades(1, nOut_ReceiveNotionals) = "ReceiveNotionals"
54        OutTrades(1, nOut_PayNotionals) = "PayNotionals"
55        OutTrades(1, nOut_NumeraireEquals) = "Numeraire=" & Numeraire

56        k = 1
57        For i = 1 To sNRows(InTrades)

58            If ChooseVector(i, 1) Then
59                TradeID = CStr(InTrades(i, nIn_TradeId))
60                VF = ProductTypeToValuationFunction(CStr(InTrades(i, nIn_ProductType)))

61                If TradesScaleFactor = 1 Then
62                    ThisScaleFactor = 1
63                Else
64                    If InTrades(i, nIn_TradeIsFrom) = tif_Future Then
65                        ThisScaleFactor = 1        'We do not scale trades sourced from the FutureTrades sheet!
66                    Else
67                        ThisScaleFactor = TradesScaleFactor
68                    End If
69                End If

70                Select Case VF
                      Case "FxForward"

71                        k = k + 1
72                        Select Case InTrades(i, nIn_BuySell)
                              Case "Buy"
73                                If InTrades(i, nIn_PrimAmt) < 0 Then Throw ("Error in trade data. Expect 'Prim Amt' to be positive when 'Buy Sell' is 'Buy', but it is negative (" & CStr(InTrades(i, nIn_PrimAmt)) & ")")
74                            Case "Sell"
75                                If InTrades(i, nIn_PrimAmt) > 0 Then Throw ("Error in trade data. Expect 'Prim Amt' to be negative when 'Buy Sell' is 'Sell', but it is positive (" & CStr(InTrades(i, nIn_PrimAmt)) & ")")
76                            Case Else
77                                Throw "Error in trade data. Expect 'Buy Sell' to be either 'Buy' or 'Sell', but it is '" & CStr(InTrades(i, nIn_BuySell)) & "'"
78                        End Select

79                        OutTrades(k, nOut_TradeID) = TradeID
80                        OutTrades(k, nOut_ValuationFunction) = VF
81                        OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)
82                        OutTrades(k, nOut_EndDate) = CDate(InTrades(i, nIn_MaturityDate) - PortfolioAgeing * 365)
83                        OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_PrimCur)
84                        OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_SecCur)
85                        OutTrades(k, nOut_ReceiveNotional) = InTrades(i, nIn_PrimAmt) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
86                        OutTrades(k, nOut_PayNotional) = InTrades(i, nIn_SecAmt) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)

87                    Case "FxSwap"
                          'This looks counter-intuitive. If we are applying compression then convert an input FXSwap into TWO output FxForwards _
                          that subsequently get compressed (together with other FxForwards) into FxForwardStrip trades. Otherwise convert an input _
                          FXSwap into a single output FxForwardStrip with two "tradelets".

88                        If Compress Then
89                            k = k + 1
90                            OutTrades(k, nOut_TradeID) = TradeID
91                            OutTrades(k, nOut_ValuationFunction) = "FxForward"
92                            OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)
93                            OutTrades(k, nOut_EndDate) = CDate(InTrades(i, nIn_SettleDate) - PortfolioAgeing * 365)
94                            OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_PrimCur)
95                            OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_SecCur)
96                            OutTrades(k, nOut_ReceiveNotional) = InTrades(i, nIn_FXPrimAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
97                            OutTrades(k, nOut_PayNotional) = InTrades(i, nIn_FXSecAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
98                            k = k + 1
99                            OutTrades(k, nOut_TradeID) = TradeID
100                           OutTrades(k, nOut_ValuationFunction) = "FxForward"
101                           OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)
102                           OutTrades(k, nOut_EndDate) = CDate(InTrades(i, nIn_MaturityDate) - PortfolioAgeing * 365)
103                           OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_PrimCur)
104                           OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_SecCur)
105                           OutTrades(k, nOut_ReceiveNotional) = InTrades(i, nIn_FXFarPrimAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
106                           OutTrades(k, nOut_PayNotional) = InTrades(i, nIn_FXFarSecAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
107                       Else
108                           k = k + 1
109                           OutTrades(k, nOut_TradeID) = TradeID
110                           OutTrades(k, nOut_ValuationFunction) = "FxForwardStrip"
111                           OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)
112                           OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_PrimCur)
113                           OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_SecCur)
114                           OutTrades(k, nOut_Dates) = Format(CDate(InTrades(i, nIn_SettleDate) - PortfolioAgeing * 365), "dd-mmm-yyyy") & ";" & _
                                  Format(CDate(InTrades(i, nIn_MaturityDate) - PortfolioAgeing * 365), "dd-mmm-yyyy")
115                           OutTrades(k, nOut_ReceiveNotionals) = CStr(InTrades(i, nIn_FXPrimAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)) & ";" & _
                                  CStr(InTrades(i, nIn_FXFarPrimAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor))
116                           OutTrades(k, nOut_PayNotionals) = CStr(InTrades(i, nIn_FXSecAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)) & ";" & _
                                  CStr(InTrades(i, nIn_FXFarSecAmount) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor))
117                       End If

118                   Case "FxOption"
119                       k = k + 1
                          'PGS 4 March 2022. Airbus have not provided explicit data on whether the option is a call or put.
                          'Working assumption - it's a call on the Prim Cur if Prim Amt is positive
                          'For Julia we represent the trade as an option on the non-numeraire currency
120                       Select Case InTrades(i, nIn_BuySell)
                              Case "Buy"
121                               LongTheOpt = True
122                           Case "Sell"
123                               LongTheOpt = False
124                           Case Else
125                               Throw "BuySell must 'Buy' or 'Sell' but it is '" & CStr(InTrades(i, nIn_BuySell)) & "'"
126                       End Select
127                       IsCallOnPrimCur = InTrades(i, nIn_PrimCur) > 0
                           
128                       If FlipTrades Then LongTheOpt = Not (LongTheOpt)
129                       If InTrades(i, nIn_PrimCur) = Numeraire Then
                              'Why? because the Julia code IsCall flag means "Is Call on NON Numeraire" but IsCallOnPrimCur in this case means "Is Call on Numeraire"
130                           IsCall = Not (IsCallOnPrimCur)
131                           Strike = Abs(InTrades(i, nIn_PrimAmt) / InTrades(i, nIn_SecAmt))
132                           Notional = Abs(InTrades(i, nIn_SecAmt)) * IIf(LongTheOpt, 1, -1) * ThisScaleFactor
133                           Ccy = InTrades(i, nIn_SecCur)
134                       ElseIf InTrades(i, nIn_SecCur) = Numeraire Then
135                           IsCall = IsCallOnPrimCur
136                           Strike = Abs(InTrades(i, nIn_SecAmt) / InTrades(i, nIn_PrimAmt))
137                           Notional = Abs(InTrades(i, nIn_PrimAmt)) * IIf(LongTheOpt, 1, -1) * ThisScaleFactor
138                           Ccy = InTrades(i, nIn_PrimCur)
139                       Else
140                           Throw "Cannot handle Fx options on cross-rates, i.e. options where neither currency is '" + Numeraire + "'"
141                       End If
142                       OutTrades(k, nOut_TradeID) = TradeID
143                       OutTrades(k, nOut_ValuationFunction) = VF
144                       OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)

145                       OutTrades(k, nOut_Currency) = Ccy
                          'For Fx Options (but not for forwards) the date in the MATURITY_DATE column is the option-expiry date, not the payment date
146                       OutTrades(k, nOut_EndDate) = CDate(AddTwoDays(CLng(InTrades(i, nIn_MaturityDate))) - PortfolioAgeing * 365)
147                       OutTrades(k, nOut_Notional) = Notional
148                       OutTrades(k, nOut_Strike) = Strike
149                       OutTrades(k, nOut_isCall) = IsCall
150                   Case "CrossCurrencySwap", "InterestRateSwap"
151                       k = k + 1
152                       Select Case InTrades(i, nIn_PrinSchedType)
                              Case "Schedule"
153                               HasSchedule = True
154                               AnySchedules = True
155                           Case "Bullet", "Custom" 'Trade 1923795 has 'Custom` but no amortisation schedule. TODO speak to Cedrick (PGS 9 March 2022)
156                               HasSchedule = False
157                           Case Else
158                               Throw "'Prin Sched Type' given as '" + CStr(InTrades(i, nIn_PrinSchedType)) + "' but allowed values are 'Bullet' or 'Schedule'"
159                       End Select
                          Dim PayIsFixed As Boolean
                          Dim ReceiveIsFixed As Boolean
160                       If VF = "InterestRateSwap" Then
161                           If InTrades(i, nIn_PayPrincipal) <> InTrades(i, nIn_RcvPrincipal) Then Throw "Pay Principal and Receive Leg Principal must be the same for ProductType = Swap"
162                           If InTrades(i, nIn_PayCcy) <> InTrades(i, nIn_RcvCcy) Then Throw "Pay Ccy and Receive Leg Principal Ccy must be the same for ProductType = Swap"
163                       End If
                      
164                       OutTrades(k, nOut_TradeID) = TradeID
165                       OutTrades(k, nOut_ValuationFunction) = VF
166                       OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CounterpartyParent)

167                       OutTrades(k, nOut_StartDate) = CDate(InTrades(i, nIn_SettleDate) - PortfolioAgeing * 365)
168                       OutTrades(k, nOut_EndDate) = CDate(InTrades(i, nIn_MaturityDate) - PortfolioAgeing * 365)
169                       OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_PayCcy)
170                       OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_RcvCcy)
171                       If HasSchedule Then
172                           OutTrades(k, nOut_ReceiveAmortNotionals) = "TBC"
173                           OutTrades(k, nOut_PayAmortNotionals) = "TBC"
174                       Else
175                           OutTrades(k, nOut_PayNotional) = -Abs(InTrades(i, nIn_PayPrincipal)) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
176                           OutTrades(k, nOut_ReceiveNotional) = Abs(InTrades(i, nIn_RcvPrincipal)) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
177                       End If
178                       OutTrades(k, nOut_PayBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_PayBDC)))
179                       OutTrades(k, nOut_ReceiveBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_RcvBDC)))
180                       OutTrades(k, nOut_PayIndex) = JuliaIndexFromAirbusTradeData(CStr(InTrades(i, nIn_PayType)), CStr(InTrades(i, nIn_PayFloatingRate)), PayIsFixed)
181                       OutTrades(k, nOut_ReceiveIndex) = JuliaIndexFromAirbusTradeData(CStr(InTrades(i, nIn_RcvType)), CStr(InTrades(i, nIn_RcvFloatingRate)), ReceiveIsFixed)
182                       If PayIsFixed Then
183                           OutTrades(k, nOut_PayCoupon) = StringWithCommaToNumber(InTrades(i, nIn_PayFixedRate)) / 100
184                           OutTrades(k, nOut_PayFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PayFreq)))
185                           OutTrades(k, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, nIn_PayDaycount)), False, True)
186                       Else
                              'Strange choice that Airbus data has only a single column for RateIndexSpread rather than one for the Pay Leg and one for the Receive Leg. Would make it not possible to book a Floating-Floating swap
187                           OutTrades(k, nOut_PayCoupon) = StringWithCommaToNumber(InTrades(i, nIn_RateIndexSpread)) / 100
188                           OutTrades(k, nOut_PayFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PayFreq)))
189                           OutTrades(k, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, nIn_PayDaycount)), True, True)
190                       End If
191                       If ReceiveIsFixed Then
192                           OutTrades(k, nOut_ReceiveCoupon) = StringWithCommaToNumber(InTrades(i, nIn_RcvFixedRate)) / 100
193                           OutTrades(k, nOut_ReceiveFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_RcvFreq)))
194                           OutTrades(k, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, nIn_RcvDaycount)), False, True)
195                       Else
196                           OutTrades(k, nOut_ReceiveCoupon) = StringWithCommaToNumber(InTrades(i, nIn_RateIndexSpread)) / 100
197                           OutTrades(k, nOut_ReceiveFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_RcvFreq)))
198                           OutTrades(k, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, nIn_RcvDaycount)), True, True)
199                       End If
200                   Case Else
201                       Throw "Unrecognised ValuationFunction: " + VF
202               End Select
203           End If
204       Next i

205       If k < sNRows(OutTrades) Then OutTrades = sSubArray(OutTrades, 1, 1, k)

          'Now deal with amortising trades...
          
206       If AnySchedules Then
              Dim AmortMap
              Dim AmortMapCol1
              Dim AmortNotionals
              Dim AmortPayRecs
              Dim AmortStartDates
              Dim AmortTradeIDs
              Dim MatchRes
              Dim PayNotionalsString As String
              Dim ReceiveNotionalsString As String
              Dim TheseNotionalDates
              Dim TheseNotionals
              Dim ThesePayNotionalDates
              Dim ThesePayNotionals
              Dim ThesePayRec
              Dim TheseReceiveNotionalDates
              Dim TheseReceiveNotionals
207           GrabAmortisationData AmortTradeIDs, AmortStartDates, AmortPayRecs, AmortNotionals, AmortMap, twb
208           AmortMapCol1 = sSubArray(AmortMap, 1, 1, , 1)
209           For i = 2 To sNRows(OutTrades)
210               If OutTrades(i, nOut_PayAmortNotionals) = "TBC" Then

211                   TradeID = OutTrades(i, nOut_TradeID)
212                   MatchRes = sMatch(CDbl(TradeID), AmortMapCol1, True)
213                   If IsNumber(MatchRes) Then
214                       TheseNotionals = sSubArray(AmortNotionals, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
215                       ThesePayRec = sSubArray(AmortPayRecs, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
216                       TheseReceiveNotionals = sMChoose(TheseNotionals, sArrayEquals(ThesePayRec, "REC"))
217                       ThesePayNotionals = sMChoose(TheseNotionals, sArrayEquals(ThesePayRec, "PAY"))
218                       TheseNotionalDates = sSubArray(AmortStartDates, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
219                       TheseReceiveNotionalDates = sMChoose(TheseNotionalDates, sArrayEquals(ThesePayRec, "REC"))
220                       ThesePayNotionalDates = sMChoose(TheseNotionalDates, sArrayEquals(ThesePayRec, "PAY"))

                          Dim FlipAndScale As Double
221                       FlipAndScale = IIf(FlipTrades, -1, 1) * TradesScaleFactor
222                       If FlipAndScale <> 1 Then
223                       TheseReceiveNotionals = sArrayMultiply(TheseReceiveNotionals, FlipAndScale)
224                       End If
225                       If FlipAndScale <> -1 Then
226                       ThesePayNotionals = sArrayMultiply(ThesePayNotionals, -FlipAndScale)
227                       End If

228                       ThesePayNotionals = InferNotionalSchedule(OutTrades(i, nOut_StartDate), OutTrades(i, nOut_EndDate), OutTrades(i, nOut_PayFrequency), CStr(OutTrades(i, nOut_PayBDC)), ThesePayNotionalDates, ThesePayNotionals, PortfolioAgeing)
229                       PayNotionalsString = sConcatenateStrings(ThesePayNotionals, ";")
230                       TheseReceiveNotionals = InferNotionalSchedule(OutTrades(i, nOut_StartDate), OutTrades(i, nOut_EndDate), OutTrades(i, nOut_ReceiveFrequency), CStr(OutTrades(i, nOut_ReceiveBDC)), TheseReceiveNotionalDates, TheseReceiveNotionals, PortfolioAgeing)
231                       ReceiveNotionalsString = sConcatenateStrings(TheseReceiveNotionals, ";")
232                       OutTrades(i, nOut_PayAmortNotionals) = PayNotionalsString
233                       OutTrades(i, nOut_ReceiveAmortNotionals) = ReceiveNotionalsString
234                   Else
235                       Throw "Trade " & TradeID & " has 'Prin Sched Type' of 'Schedule' but no amortisation data is found in the amortisation file"
236                   End If
237               End If
238           Next i
239       End If

240       If Compress Then
              Dim numCompressedTrades As Long
241           OutTrades = CompressJuliaFxForwards(OutTrades, numCompressedTrades)
242           OutTrades = CompressJuliaFxOptions(OutTrades, numCompressedTrades + 1)
243       End If
          'Cache results...
244       PreviousCheckSum = ThisCheckSum
245       PreviousResult = OutTrades
246       PreviousTC = TC
247       GetTradesInJuliaFormat = OutTrades

248       Exit Function
ErrHandler:

249       ErrString = "#GetTradesInJuliaFormat (line " & CStr(Erl) & "): " & Err.Description
250       If TradeID <> "" Then
251           ErrString = ErrString & " (TradeID = " & TradeID & ")!"
252       Else
253           ErrString = ErrString & "!"
254       End If
255       GetTradesInJuliaFormat = ErrString
256       PreviousResult = Empty
257       PreviousCheckSum = ""
258       PreviousNumTrades = 0
259       Throw ErrString
End Function

Function ProductTypeToValuationFunction(ProductType As String)
1         On Error GoTo ErrHandler
2         Select Case ProductType
              Case "FXForward", "FXNDF"
3                 ProductTypeToValuationFunction = "FxForward"
4             Case "FXSwap"
                  'In fact there is no such ValuationFunction in the Julia code...
5                 ProductTypeToValuationFunction = "FxSwap"
6             Case "FXOption"
7                 ProductTypeToValuationFunction = "FxOption"
8             Case "Swap"
9                 ProductTypeToValuationFunction = "InterestRateSwap"
10            Case "XCCySwap"
11                ProductTypeToValuationFunction = "CrossCurrencySwap"
12            Case Else
13                Throw "Unrecognised value in 'Product Type' column: '" + ProductType + "'"
14        End Select
15        Exit Function
ErrHandler:
16        Throw "#ProductTypeToValuationFunction (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function JuliaIndexFromAirbusTradeData(LegType As String, FloatingRate As String, ByRef isFixed As Boolean)

1         On Error GoTo ErrHandler
2         If LegType = "Fixed" Then
3             JuliaIndexFromAirbusTradeData = "Fixed"
4             isFixed = True
5         ElseIf LegType = "Float" Then
6             isFixed = False
7             If InStr(FloatingRate, "IBOR") > 0 Then
8                 JuliaIndexFromAirbusTradeData = "Libor"
9             Else
                  'We have not seen any Airbus trade data for OIS swaps
10                JuliaIndexFromAirbusTradeData = "OIS"
11            End If
12        Else
13            Throw "LegType must be either 'Fixed' or 'Float'"
14        End If

15        Exit Function
ErrHandler:
16        Throw "#JuliaIndexFromAirbusTradeData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompressJuliaFxOptions
' Author    : Philip Swannell
' Date      : 26-Dec-2016
' Purpose   : replace all FxOption trades in a portfolio with FxOptionStrip trades
' -----------------------------------------------------------------------------------------------------------------------
Function CompressJuliaFxOptions(TheTrades, FirstIndex As Long)
          Dim ChooseVector As Variant
          Dim CountRepeatsRet As Variant
          Dim Headers As Variant
          Dim i As Long
          Dim j As Long
          Dim nIn_Counterparty As Variant
          Dim nIn_Currency As Variant
          Dim nIn_Dates As Variant
          Dim nIn_EndDate As Variant
          Dim nIn_isCall As Variant
          Dim nIn_Notional As Variant
          Dim nIn_Notionals As Variant
          Dim nIn_Strike As Variant
          Dim nIn_Strikes As Variant
          Dim nIn_TradeId As Variant
          Dim nIn_ValuationFunction As Variant
          Dim NumFxOptions As Long
          Dim numFxOptionStrips As Long
          Dim numNotFxOptions As Long
          Dim NumTrades As Long
          Dim SortResult
          Dim TheFxOptions As Variant

1         On Error GoTo ErrHandler
2         Headers = sArrayTranspose(sSubArray(TheTrades, 1, 1, 1))
3         nIn_TradeId = sMatch("TradeID", Headers): If Not IsNumber(nIn_TradeId) Then Throw ("Cannot find 'TradeID' in row 2 of TheTrades")
4         nIn_ValuationFunction = sMatch("ValuationFunction", Headers): If Not IsNumber(nIn_ValuationFunction) Then Throw ("Cannot find 'ValuationFunction' in row 2 of TheTrades")
5         nIn_Counterparty = sMatch("Counterparty", Headers): If Not IsNumber(nIn_Counterparty) Then Throw ("Cannot find 'Counterparty' in row 2 of TheTrades")
6         nIn_EndDate = sMatch("EndDate", Headers): If Not IsNumber(nIn_EndDate) Then Throw ("Cannot find 'EndDate' in row 2 of TheTrades")
7         nIn_Currency = sMatch("Currency", Headers): If Not IsNumber(nIn_Currency) Then Throw ("Cannot find 'Currency' in row 2 of TheTrades")
8         nIn_Notional = sMatch("Notional", Headers): If Not IsNumber(nIn_Notional) Then Throw ("Cannot find 'Notional' in row 2 of TheTrades")
9         nIn_Strike = sMatch("Strike", Headers): If Not IsNumber(nIn_Strike) Then Throw ("Cannot find 'Strike' in row 2 of TheTrades")
10        nIn_isCall = sMatch("isCall", Headers): If Not IsNumber(nIn_isCall) Then Throw ("Cannot find 'isCall' in row 2 of TheTrades")
11        nIn_Dates = sMatch("Dates", Headers)
12        nIn_Notionals = sMatch("Notionals", Headers)
13        nIn_Strikes = sMatch("Strikes", Headers)

14        NumTrades = sNRows(TheTrades) - 1
15        ChooseVector = sArrayEquals(sSubArray(TheTrades, 1, nIn_ValuationFunction, , 1), "FxOption")
16        NumFxOptions = sArrayCount(ChooseVector)
17        numNotFxOptions = NumTrades - NumFxOptions

18        If NumFxOptions = 0 Then
19            CompressJuliaFxOptions = TheTrades
20            Exit Function
21        End If

22        TheFxOptions = sMChoose(TheTrades, ChooseVector)

23        SortResult = sArrayRange(sSubArray(TheFxOptions, 1, nIn_Counterparty, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_Currency, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_isCall, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_EndDate, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_Strike, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_Notional, , 1))

          'Could use SortMerge here to mash together trades with the same Counterparty, Currency, isCall and Strike but in practice that's unlikely to lead to much (any?) reduction in portfolio size

24        SortResult = sSortedArray(SortResult, 1, 2, 3, True, True, True, False)
25        CountRepeatsRet = sCountRepeats(sRowConcatenateStrings(sSubArray(SortResult, 1, 1, , 3)), "FH")
26        numFxOptionStrips = sNRows(CountRepeatsRet)

          Dim ChooseVector2
          Dim ExtraCol
          Dim ReturnArray
          Dim TheseDates
          Dim ThisCcy
          Dim ThisCpty

          'Construct the return including space for the FxOptionStrip trades
27        ChooseVector2 = sArrayNot(ChooseVector)
28        ReturnArray = sMChoose(TheTrades, ChooseVector2)
29        If Not IsNumber(nIn_Dates) Then
30            ExtraCol = "Dates"
31            If numNotFxOptions > 0 Then
32                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, numNotFxOptions, 1))
33            End If
34            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
35            nIn_Dates = sNCols(ReturnArray)
36        End If
37        If Not IsNumber(nIn_Notionals) Then
38            ExtraCol = "Notionals"
39            If numNotFxOptions > 0 Then
40                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, numNotFxOptions, 1))
41            End If
42            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
43            nIn_Notionals = sNCols(ReturnArray)
44        End If
45        If Not IsNumber(nIn_Strikes) Then
46            ExtraCol = "Strikes"
47            If numNotFxOptions > 0 Then
48                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, numNotFxOptions, 1))
49            End If
50            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
51            nIn_Strikes = sNCols(ReturnArray)
52        End If

          Dim FxOptionStripTrades

53        FxOptionStripTrades = sReshape(Empty, numFxOptionStrips, sNCols(ReturnArray))

68        For i = 1 To numFxOptionStrips
              Dim From As Long
              Dim HowMany As Long
              Dim TheseNotionals As Variant
              Dim TheseStrikes
              Dim ThisIsCall As Boolean

69            From = CountRepeatsRet(i, 1)
70            HowMany = CountRepeatsRet(i, 2)

71            ThisCpty = SortResult(From, 1)
72            ThisCcy = SortResult(From, 2)
73            ThisIsCall = SortResult(From, 3)
74            TheseDates = sSubArray(SortResult, From, 4, HowMany, 1)
75            TheseNotionals = sSubArray(SortResult, From, 6, HowMany, 1)
76            TheseStrikes = sSubArray(SortResult, From, 5, HowMany, 1)
77            For j = 1 To HowMany
78                TheseDates(j, 1) = Format(TheseDates(j, 1), "dd-mmm-yyyy")
79            Next j
80            FxOptionStripTrades(i, nIn_TradeId) = "Compression" & CStr(FirstIndex - 1 + i)
81            FxOptionStripTrades(i, nIn_ValuationFunction) = "FxOptionStrip"
82            FxOptionStripTrades(i, nIn_Counterparty) = ThisCpty
83            FxOptionStripTrades(i, nIn_Currency) = ThisCcy
84            FxOptionStripTrades(i, nIn_isCall) = ThisIsCall
85            FxOptionStripTrades(i, nIn_Dates) = sConcatenateStrings(TheseDates, ";")
86            FxOptionStripTrades(i, nIn_Notionals) = sConcatenateStrings(TheseNotionals, ";")
87            FxOptionStripTrades(i, nIn_Strikes) = sConcatenateStrings(TheseStrikes, ";")
88        Next i

89        CompressJuliaFxOptions = sArrayStack(ReturnArray, FxOptionStripTrades)

90        Exit Function
ErrHandler:
91        Throw "#CompressJuliaFxOptions (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompressJuliaFxForwards
' Author    : Philip Swannell
' Date      : 01-Jan-2017
' Purpose   : Compress FxForward trades to FxForwardStrip trades rather than to FixedCashflows
'             trades. Advantage is that Capital calculations based on trade notionals are then
'             possible
' -----------------------------------------------------------------------------------------------------------------------
Function CompressJuliaFxForwards(TheTrades, ByRef numCompressedTrades As Long)
          Dim ChooseVector As Variant
          Dim CountRepeatsRet As Variant
          Dim Headers As Variant
          Dim i As Long
          Dim j As Long
          Dim nIn_Counterparty As Variant
          Dim nIn_Dates As Variant
          Dim nIn_EndDate As Variant
          Dim nIn_PayCurrency As Variant
          Dim nIn_PayNotional As Variant
          Dim nIn_PayNotionals As Variant
          Dim nIn_ReceiveCurrency As Variant
          Dim nIn_ReceiveNotional As Variant
          Dim nIn_ReceiveNotionals As Variant
          Dim nIn_TradeId As Variant
          Dim nIn_ValuationFunction As Variant
          Dim NumFxForwards As Long
          Dim NumFxForwardStrips As Long
          Dim NumNotFxForwards As Long
          Dim NumTrades As Long
          Dim SortResult
          Dim TheFxForwards As Variant

1         On Error GoTo ErrHandler
2         Headers = sArrayTranspose(sSubArray(TheTrades, 1, 1, 1))
3         nIn_TradeId = sMatch("TradeID", Headers): If Not IsNumber(nIn_TradeId) Then Throw ("Cannot find 'TradeID' in row 2 of TheTrades")
4         nIn_ValuationFunction = sMatch("ValuationFunction", Headers): If Not IsNumber(nIn_ValuationFunction) Then Throw ("Cannot find 'ValuationFunction' in row 2 of TheTrades")
5         nIn_Counterparty = sMatch("Counterparty", Headers): If Not IsNumber(nIn_Counterparty) Then Throw ("Cannot find 'Counterparty' in row 2 of TheTrades")
6         nIn_EndDate = sMatch("EndDate", Headers): If Not IsNumber(nIn_EndDate) Then Throw ("Cannot find 'EndDate' in row 2 of TheTrades")
7         nIn_ReceiveCurrency = sMatch("ReceiveCurrency", Headers): If Not IsNumber(nIn_ReceiveCurrency) Then Throw ("Cannot find 'ReceiveCurrency' in row 2 of TheTrades")
8         nIn_PayCurrency = sMatch("PayCurrency", Headers): If Not IsNumber(nIn_PayCurrency) Then Throw ("Cannot find 'PayCurrency' in row 2 of TheTrades")
9         nIn_ReceiveNotional = sMatch("ReceiveNotional", Headers): If Not IsNumber(nIn_ReceiveNotional) Then Throw ("Cannot find 'ReceiveNotional' in row 2 of TheTrades")
10        nIn_PayNotional = sMatch("PayNotional", Headers): If Not IsNumber(nIn_PayNotional) Then Throw ("Cannot find 'PayNotional' in row 2 of TheTrades")
11        nIn_Dates = sMatch("Dates", Headers)
12        nIn_ReceiveNotionals = sMatch("ReceiveNotionals", Headers)
13        nIn_PayNotionals = sMatch("PayNotionals", Headers)

14        NumTrades = sNRows(TheTrades) - 1
15        ChooseVector = sArrayEquals(sSubArray(TheTrades, 1, nIn_ValuationFunction, , 1), "FxForward")
16        NumFxForwards = sArrayCount(ChooseVector)
17        NumNotFxForwards = NumTrades - NumFxForwards

18        If NumFxForwards = 0 Then
19            CompressJuliaFxForwards = TheTrades
20            Exit Function
21        End If

22        TheFxForwards = sMChoose(TheTrades, ChooseVector)

23        SortResult = sArrayRange(sSubArray(TheFxForwards, 1, nIn_Counterparty, , 1), _
                                   sSubArray(TheFxForwards, 1, nIn_ReceiveCurrency, , 1), _
                                   sSubArray(TheFxForwards, 1, nIn_PayCurrency, , 1), _
                                   sSubArray(TheFxForwards, 1, nIn_EndDate, , 1), _
                                   sSubArray(TheFxForwards, 1, nIn_ReceiveNotional, , 1), _
                                   sSubArray(TheFxForwards, 1, nIn_PayNotional, , 1))

24        SortResult = sSortedArray(SortResult, 1, 2, 3, True, True, True, False)
25        CountRepeatsRet = sCountRepeats(sRowConcatenateStrings(sSubArray(SortResult, 1, 1, , 3)), "FH")
26        NumFxForwardStrips = sNRows(CountRepeatsRet)

          Dim ChooseVector2
          Dim ExtraCol
          Dim ReturnArray
          Dim TheseDates
          Dim ThisCpty
          Dim ThisReceiveCurrency

          'Construct the return including space for the FxForwardStrip trades
27        ChooseVector2 = sArrayNot(ChooseVector)
28        ReturnArray = sMChoose(TheTrades, ChooseVector2)
29        If Not IsNumber(nIn_Dates) Then
30            ExtraCol = "Dates"
31            If NumNotFxForwards > 0 Then
32                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, NumNotFxForwards, 1))
33            End If
34            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
35            nIn_Dates = sNCols(ReturnArray)
36        End If
37        If Not IsNumber(nIn_ReceiveNotionals) Then
38            ExtraCol = "ReceiveNotionals"
39            If NumNotFxForwards > 0 Then
40                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, NumNotFxForwards, 1))
41            End If
42            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
43            nIn_ReceiveNotionals = sNCols(ReturnArray)
44        End If
45        If Not IsNumber(nIn_PayNotionals) Then
46            ExtraCol = "PayNotionals"
47            If NumNotFxForwards > 0 Then
48                ExtraCol = sArrayStack(ExtraCol, sReshape(Empty, NumNotFxForwards, 1))
49            End If
50            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
51            nIn_PayNotionals = sNCols(ReturnArray)
52        End If

          Dim FxForwardStripTrades

53        FxForwardStripTrades = sReshape(Empty, NumFxForwardStrips, sNCols(ReturnArray))

68        For i = 1 To NumFxForwardStrips
              Dim From As Long
              Dim HowMany As Long
              Dim ThesePayNotionals
              Dim TheseReceiveNotionals As Variant
              Dim ThisPayCurrency As String

69            From = CountRepeatsRet(i, 1)
70            HowMany = CountRepeatsRet(i, 2)

71            ThisCpty = SortResult(From, 1)
72            ThisReceiveCurrency = SortResult(From, 2)
73            ThisPayCurrency = SortResult(From, 3)
74            TheseDates = sSubArray(SortResult, From, 4, HowMany, 1)
75            TheseReceiveNotionals = sSubArray(SortResult, From, 5, HowMany, 1)
76            ThesePayNotionals = sSubArray(SortResult, From, 6, HowMany, 1)
77            For j = 1 To HowMany
78                TheseDates(j, 1) = Format(TheseDates(j, 1), "dd-mmm-yyyy")
79            Next j
80            FxForwardStripTrades(i, nIn_TradeId) = "Compression" & CStr(i)
81            FxForwardStripTrades(i, nIn_ValuationFunction) = "FxForwardStrip"
82            FxForwardStripTrades(i, nIn_Counterparty) = ThisCpty
83            FxForwardStripTrades(i, nIn_ReceiveCurrency) = ThisReceiveCurrency
84            FxForwardStripTrades(i, nIn_PayCurrency) = ThisPayCurrency
85            FxForwardStripTrades(i, nIn_Dates) = sConcatenateStrings(TheseDates, ";")
86            FxForwardStripTrades(i, nIn_ReceiveNotionals) = sConcatenateStrings(TheseReceiveNotionals, ";")
87            FxForwardStripTrades(i, nIn_PayNotionals) = sConcatenateStrings(ThesePayNotionals, ";")
88        Next i

89        numCompressedTrades = NumFxForwardStrips
90        CompressJuliaFxForwards = sArrayStack(ReturnArray, FxForwardStripTrades)

91        Exit Function
ErrHandler:
92        Throw "#CompressJuliaFxForwards (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

