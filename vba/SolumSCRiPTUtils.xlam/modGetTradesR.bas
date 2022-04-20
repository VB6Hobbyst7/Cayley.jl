Attribute VB_Name = "modGetTradesR"
Option Explicit

Public Const MT_HW = "Multi-Currency Hull-White"
Public Const MT_LnFx = "Log-Normal Fx Only"
'Names of the sheets in the Market Data Workbook
Public Const SN_RatesTrades = "~ IRD BOOK"
Public Const SN_FxTrades = "~ FX BOOK"
Public Const SN_Amortisation = "~ IRD AMORTISATION"
'Better sheet names for Cayley2022
Public Const SN_RatesTrades2 = "Rates"
Public Const SN_FxTrades2 = "Fx"
Public Const SN_Amortisation2 = "Amortisation"

Public Const SN_Lines = "Summary"

Public Type TradeCount
    NumExcluded As Long
    NumIncluded As Long
    Total As Long
End Type

'Enumeration of the TradeIsFrom that can be returned by GetColumnsFromTradesWorkbook
Public Enum TradeIsFrom
    tif_Fx = 0
    tif_Rates = 1
    tif_Future = 2
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetTradesInRFormat
' Author    : Philip Swannell
' Date      : 13-Jul-2016
' Purpose   : Takes trades from the trades workbook and puts then in the format needed
'             for R. Also applies PortfolioAgeing by amending the trades' maturity dates.
' -----------------------------------------------------------------------------------------------------------------------
Function GetTradesInRFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
                            PortfolioAgeing As Double, FlipTrades As Boolean, Numeraire As String, _
                            WithFxTrades As Boolean, WithRatesTrades As Boolean, TradesScaleFactor As Double, _
                            CurrenciesToInclude As String, ModelName As String, Compress As Boolean, _
                            TC As TradeCount, twb As Workbook, fts As Worksheet)

1         On Error GoTo ErrHandler
          'Frequently call with same arguments as last time, so use cacheing. _
           Note that NumTrades should not be part of the CheckSum.
          'Cacheing is valid as long as the trades in the trades workbook and the _
           FutureTrades sheet have not changed since the previous call. So we have _
           method FlushStatics which gets called after that might have happened
          Dim ThisCheckSum As String
          Static PreviousCheckSum As String
          Static PreviousResult As Variant
          Static PreviousNumTrades As Long
          Static PreviousTC As TradeCount

2         ThisCheckSum = CStr(FilterBy1) & "," & CStr(Filter1Value) & "," & CStr(FilterBy2) & "," & _
                         CStr(Filter2Value) & "," & CStr(IncludeFutureTrades) & "," & CStr(PortfolioAgeing) & "," & _
                         CStr(FlipTrades) & "," & CStr(Numeraire) & "," & CStr(WithFxTrades) & "," & _
                         CStr(WithRatesTrades) & "," & CStr(TradesScaleFactor) & "," & CStr(CurrenciesToInclude) & "," & _
                         CStr(ModelName) & "," & CStr(Compress) & "," & twb.FullName
3         If ThisCheckSum = PreviousCheckSum Then
4             If Not (IsEmpty(PreviousResult)) Then
5                 GetTradesInRFormat = PreviousResult
6                 TC = PreviousTC
7                 Exit Function
8             End If
9         End If

          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double
          Dim t5 As Double
10        t1 = sElapsedTime()
          Dim Headers() As String
11        ReDim Headers(1 To 2, 1 To 9)
          Dim AnyIRDs As Boolean
          Dim Ccy As String
          Dim ChooseVector As Variant
          Dim ErrString As String
          Dim i As Long
          Dim InTrades As Variant
          Dim IsCall As Boolean
          Dim IsCallOnCCY1 As Boolean
          Dim k As Long
          Dim LongTheOpt As Boolean
          Dim Notional As Double
          Dim NumOutTrades As Long
          Dim OutTrades() As Variant
          Dim Strike As Double
          Dim ThisScaleFactor As Double
          Dim TradeID As String
          Dim VF As String
          Dim AnchorDate As Date

          'These variables must be in synch with definition of first two lines of outTrades
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
          '    Const nOut_CashflowAmounts = 14
          '    Const nOut_CashflowDates = 15
          Const nOut_Coupon = 16
          Const nOut_FixedAmortNotionals = 17
          Const nOut_FixedBDC = 18
          Const nOut_FixedDCT = 19
          Const nOut_FixedFrequency = 20
          Const nOut_FloatingAmortNotionals = 21
          Const nOut_FloatingBDC = 22
          Const nOut_FloatingDCT = 23
          Const nOut_FloatingFrequency = 24
          Const nOut_Margin = 25
          Const nOut_PayAmortNotionals = 26
          Const nOut_PayBDC = 27
          Const nOut_PayCoupon = 28
          Const nOut_PayDCT = 29
          Const nOut_PayFrequency = 30
          Const nOut_PayIsFixed = 31
          Const nOut_ReceiveAmortNotionals = 32
          Const nOut_ReceiveBDC = 33
          Const nOut_ReceiveCoupon = 34
          Const nOut_ReceiveDCT = 35
          Const nOut_ReceiveFrequency = 36
          Const nOut_ReceiveIsFixed = 37

          'Changing the headers we look for? Then amend method CheckTradesWorkbook
          Const HeaderNames = "OP_FINANCE,DEAL_TYPE,CPTY_PARENT,VALUE_DATE,MATURITY_DATE,CCY1,CCY2,FWD_1,FWD_2,BASIS_PAY,BASIS_REC,CCY_PAY,CCY_REC,INDEX_PAY,INDEX_REC,NOMINAL_PAY,NOMINAL_REC,SPREAD_PAY,SPREAD_REC,TradeIsFrom,PAY_COUPON_FREQ,REC_COUPON_FREQ,PAY_COUPON_DTROLL,REC_COUPON_DTROLL"
          'These variables must be in synch with the constant HeaderNames
          Const nIn_OP_FINANCE = 1
          Const nIn_DEAL_TYPE = 2
          Const nIn_CPTY_PARENT = 3
          Const nIn_VALUE_DATE = 4
          Const nIn_MATURITY_DATE = 5
          Const nIn_CCY1 = 6
          Const nIn_CCY2 = 7
          Const nIn_FWD_1 = 8
          Const nIn_FWD_2 = 9
          Const nIn_BASIS_PAY = 10
          Const nIn_BASIS_REC = 11
          Const nIn_CCY_PAY = 12
          Const nIn_CCY_REC = 13
          Const nIn_INDEX_PAY = 14
          Const nIn_INDEX_REC = 15
          Const nIn_NOMINAL_PAY = 16
          Const nIn_NOMINAL_REC = 17
          Const nIn_SPREAD_PAY = 18
          Const nIn_SPREAD_REC = 19
          Const nIn_TradeIsFrom = 20
          Const nIn_PAY_COUPON_FREQ = 21
          Const nIn_REC_COUPON_FREQ = 22
          Const nIn_PAY_COUPON_DTROLL = 23
          Const nIn_REC_COUPON_DTROLL = 24

12        On Error GoTo ErrHandler
13        If PortfolioAgeing < 0 Then Throw "PortfolioAgeing must be zero or positive"

14        AnchorDate = MyAnchorDate(ModelName)

15        t2 = sElapsedTime()
16        ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, PortfolioAgeing, False, WithFxTrades, WithRatesTrades, CurrenciesToInclude, TC, twb, fts, AnchorDate)
17        t3 = sElapsedTime
18        InTrades = GetColumnsFromTradesWorkbook(HeaderNames, IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
19        t4 = sElapsedTime
20        NumOutTrades = sArrayCount(ChooseVector)

21        ReDim OutTrades(1 To NumOutTrades + 2, 1 To 37)
22        OutTrades(1, 1) = "CHAR": OutTrades(1, 2) = "CHAR": OutTrades(1, 3) = "CHAR": OutTrades(1, 4) = "DATESTR": OutTrades(1, 5) = "DATESTR": OutTrades(1, 6) = "CHAR": OutTrades(1, 7) = "CHAR": OutTrades(1, 8) = "CHAR": OutTrades(1, 9) = "DOUBLE": OutTrades(1, 10) = "DOUBLE"
23        OutTrades(1, 11) = "DOUBLE": OutTrades(1, 12) = "DOUBLE": OutTrades(1, 13) = "BOOL": OutTrades(1, 14) = "CHAR": OutTrades(1, 15) = "CHAR": OutTrades(1, 16) = "DOUBLE": OutTrades(1, 17) = "CHAR": OutTrades(1, 18) = "CHAR": OutTrades(1, 19) = "CHAR": OutTrades(1, 20) = "INT"
24        OutTrades(1, 21) = "CHAR": OutTrades(1, 22) = "CHAR": OutTrades(1, 23) = "CHAR": OutTrades(1, 24) = "INT": OutTrades(1, 25) = "DOUBLE": OutTrades(1, 26) = "CHAR": OutTrades(1, 27) = "CHAR": OutTrades(1, 28) = "DOUBLE": OutTrades(1, 29) = "CHAR": OutTrades(1, 30) = "INT"
25        OutTrades(1, 31) = "BOOL": OutTrades(1, 32) = "CHAR": OutTrades(1, 33) = "CHAR": OutTrades(1, 34) = "DOUBLE": OutTrades(1, 35) = "CHAR": OutTrades(1, 36) = "INT": OutTrades(1, 37) = "BOOL"
26        OutTrades(2, 1) = "TradeID": OutTrades(2, 2) = "ValuationFunction": OutTrades(2, 3) = "Counterparty": OutTrades(2, 4) = "StartDate": OutTrades(2, 5) = "EndDate": OutTrades(2, 6) = "Currency": OutTrades(2, 7) = "ReceiveCurrency": OutTrades(2, 8) = "PayCurrency": OutTrades(2, 9) = "Notional": OutTrades(2, 10) = "ReceiveNotional"
27        OutTrades(2, 11) = "PayNotional": OutTrades(2, 12) = "Strike": OutTrades(2, 13) = "IsCall": OutTrades(2, 14) = "CashflowAmounts": OutTrades(2, 15) = "CashflowDates": OutTrades(2, 16) = "Coupon": OutTrades(2, 17) = "FixedAmortNotionals": OutTrades(2, 18) = "FixedBDC": OutTrades(2, 19) = "FixedDCT": OutTrades(2, 20) = "FixedFrequency"
28        OutTrades(2, 21) = "FloatingAmortNotionals": OutTrades(2, 22) = "FloatingBDC": OutTrades(2, 23) = "FloatingDCT": OutTrades(2, 24) = "FloatingFrequency": OutTrades(2, 25) = "Margin": OutTrades(2, 26) = "PayAmortNotionals": OutTrades(2, 27) = "PayBDC": OutTrades(2, 28) = "PayCoupon": OutTrades(2, 29) = "PayDCT": OutTrades(2, 30) = "PayFrequency"
29        OutTrades(2, 31) = "PayIsFixed": OutTrades(2, 32) = "ReceiveAmortNotionals": OutTrades(2, 33) = "ReceiveBDC": OutTrades(2, 34) = "ReceiveCoupon": OutTrades(2, 35) = "ReceiveDCT": OutTrades(2, 36) = "ReceiveFrequency": OutTrades(2, 37) = "ReceiveIsFixed"

30        k = 2
31        For i = 1 To sNRows(InTrades)

32            If ChooseVector(i, 1) Then
33                TradeID = CStr(InTrades(i, nIn_OP_FINANCE))
34                VF = DealTypeToValuationFunction(CStr(InTrades(i, nIn_DEAL_TYPE)))

35                If TradesScaleFactor = 1 Then
36                    ThisScaleFactor = 1
37                Else
38                    If InTrades(i, nIn_TradeIsFrom) = tif_Future Then
39                        ThisScaleFactor = 1        'We do not scale trades sourced from the FutureTrades sheet!
40                    Else
41                        ThisScaleFactor = TradesScaleFactor
42                    End If
43                End If

44                Select Case VF
                  Case "FxForward"

45                    k = k + 1
46                    OutTrades(k, nOut_TradeID) = TradeID
47                    OutTrades(k, nOut_ValuationFunction) = VF
48                    OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CPTY_PARENT)

49                    OutTrades(k, nOut_StartDate) = InTrades(i, nIn_VALUE_DATE)
50                    OutTrades(k, nOut_EndDate) = InTrades(i, nIn_MATURITY_DATE) - PortfolioAgeing * 365
51                    OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_CCY1)
52                    OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_CCY2)
53                    OutTrades(k, nOut_ReceiveNotional) = InTrades(i, nIn_FWD_1) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
54                    OutTrades(k, nOut_PayNotional) = InTrades(i, nIn_FWD_2) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)

55                Case "FxOption"
56                    k = k + 1
                      'For R we represent the trade as an option on the non-numeraire currency
57                    Select Case CStr(InTrades(i, nIn_DEAL_TYPE))
                      Case "CALLbuy VANILLA"
58                        LongTheOpt = True
59                        IsCallOnCCY1 = True
60                    Case "CALLsell VANILLA"
61                        LongTheOpt = False
62                        IsCallOnCCY1 = True
63                    Case "PUTbuy VANILLA"
64                        LongTheOpt = True
65                        IsCallOnCCY1 = False
66                    Case "PUTsell VANILLA"
67                        LongTheOpt = False
68                        IsCallOnCCY1 = False
69                    Case Else
70                        Throw "Unrecognised DEAL_TYPE"
71                    End Select
72                    If FlipTrades Then LongTheOpt = Not (LongTheOpt)
73                    If InTrades(i, nIn_CCY1) = Numeraire Then
                          'Why? because the R code IsCall flag means "Is Call on NON Numeraire" but IsCallOnCCY1 in this case means "Is Call on Numeraire"
74                        IsCall = Not (IsCallOnCCY1)
75                        Strike = Abs(InTrades(i, nIn_FWD_1) / InTrades(i, nIn_FWD_2))
76                        Notional = Abs(InTrades(i, nIn_FWD_2)) * IIf(LongTheOpt, 1, -1) * ThisScaleFactor
77                        Ccy = InTrades(i, nIn_CCY2)
78                    ElseIf InTrades(i, nIn_CCY2) = Numeraire Then
79                        IsCall = IsCallOnCCY1
80                        Strike = Abs(InTrades(i, nIn_FWD_2) / InTrades(i, nIn_FWD_1))
81                        Notional = Abs(InTrades(i, nIn_FWD_1)) * IIf(LongTheOpt, 1, -1) * ThisScaleFactor
82                        Ccy = InTrades(i, nIn_CCY1)
83                    Else
84                        Throw "Cannot handle Fx options on cross-rates, i.e. options where neither currency is '" + Numeraire + "'"
85                    End If
86                    OutTrades(k, nOut_TradeID) = TradeID
87                    OutTrades(k, nOut_ValuationFunction) = VF
88                    OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CPTY_PARENT)

89                    OutTrades(k, nOut_Currency) = Ccy
                      'For Fx Options (but not for forwards) the date in the MATURITY_DATE column is the option-expiry date, not the payment date
90                    OutTrades(k, nOut_EndDate) = AddTwoDays(CLng(InTrades(i, nIn_MATURITY_DATE))) - PortfolioAgeing * 365
91                    OutTrades(k, nOut_Notional) = Notional
92                    OutTrades(k, nOut_Strike) = Strike
93                    OutTrades(k, nOut_isCall) = IsCall
94                Case "InterestRateSwap"
95                    k = k + 1
96                    AnyIRDs = True
97                    If InTrades(i, nIn_NOMINAL_PAY) <> InTrades(i, nIn_NOMINAL_REC) Then Throw "NOMINAL_PAY and NOMINAL_REC must be the same for DEAL_TYPE = Swap"
98                    If InTrades(i, nIn_CCY_PAY) <> InTrades(i, nIn_CCY_REC) Then Throw "CCY_PAY and CCY_REC must be the same for DEAL_TYPE = Swap"
99                    OutTrades(k, nOut_TradeID) = TradeID
100                   OutTrades(k, nOut_ValuationFunction) = VF
101                   OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CPTY_PARENT)
102                   OutTrades(k, nOut_StartDate) = InTrades(i, nIn_VALUE_DATE)
103                   OutTrades(k, nOut_EndDate) = InTrades(i, nIn_MATURITY_DATE) - PortfolioAgeing * 365
104                   OutTrades(k, nOut_Currency) = InTrades(i, nIn_CCY_REC)
105                   If InTrades(i, nIn_SPREAD_PAY) = "FIX" And InTrades(i, nIn_SPREAD_REC) <> "FIX" Then
                          'Airbus pays fixed (iff FlipTrades = False)
                          'Field ReceiveIsFixed is not used by the R code (for InterestRateSwap), but it is referenced later in this method when dealing with amortising notionals
106                       OutTrades(k, nOut_ReceiveIsFixed) = False
107                       OutTrades(k, nOut_Notional) = InTrades(i, nIn_NOMINAL_PAY) * IIf(FlipTrades, ThisScaleFactor, -ThisScaleFactor)
108                       OutTrades(k, nOut_Coupon) = StringWithCommaToNumber(InTrades(i, nIn_INDEX_PAY)) / 100
109                       OutTrades(k, nOut_Margin) = StringWithCommaToNumber(InTrades(i, nIn_SPREAD_REC)) / 100
110                       OutTrades(k, nOut_FixedBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_PAY_COUPON_DTROLL)))
111                       OutTrades(k, nOut_FloatingBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_REC_COUPON_DTROLL)))
112                       OutTrades(k, nOut_FixedDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_PAY)), False, True)
113                       OutTrades(k, nOut_FloatingDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_REC)), True, True)
114                       OutTrades(k, nOut_FixedFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PAY_COUPON_FREQ)))
115                       OutTrades(k, nOut_FloatingFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_REC_COUPON_FREQ)))
116                   ElseIf InTrades(i, nIn_SPREAD_PAY) <> "FIX" And InTrades(i, nIn_SPREAD_REC) = "FIX" Then
                          'Airbus receives fixed (iff FlipTrades = False)
                          'Field ReceiveIsFixed is not used by the R code (for InterestRateSwap), but it is referenced later in this method when dealing with amortising notionals
117                       OutTrades(k, nOut_ReceiveIsFixed) = True
118                       OutTrades(k, nOut_Notional) = InTrades(i, nIn_NOMINAL_PAY) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
119                       OutTrades(k, nOut_Coupon) = StringWithCommaToNumber(InTrades(i, nIn_INDEX_REC)) / 100
120                       OutTrades(k, nOut_Margin) = StringWithCommaToNumber(InTrades(i, nIn_SPREAD_PAY)) / 100
121                       OutTrades(k, nOut_FixedBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_REC_COUPON_DTROLL)))
122                       OutTrades(k, nOut_FloatingBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_PAY_COUPON_DTROLL)))
123                       OutTrades(k, nOut_FixedDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_REC)), False, True)
124                       OutTrades(k, nOut_FloatingDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_PAY)), True, True)
125                       OutTrades(k, nOut_FixedFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_REC_COUPON_FREQ)))
126                       OutTrades(k, nOut_FloatingFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PAY_COUPON_FREQ)))
127                   Else
128                       Throw "For DEAL_TYPE = 'Swap' exactly one of SPREAD_PAY and SPREAD_REC must be 'FIX'"
129                   End If
130               Case "CrossCurrencySwap"
131                   k = k + 1
132                   AnyIRDs = True
                      Dim PayIsFixed As Boolean
                      Dim ReceiveIsFixed As Boolean
133                   OutTrades(k, nOut_TradeID) = TradeID
134                   OutTrades(k, nOut_ValuationFunction) = VF
135                   OutTrades(k, nOut_Counterparty) = InTrades(i, nIn_CPTY_PARENT)

136                   OutTrades(k, nOut_StartDate) = InTrades(i, nIn_VALUE_DATE)
137                   OutTrades(k, nOut_EndDate) = InTrades(i, nIn_MATURITY_DATE) - PortfolioAgeing * 365
138                   OutTrades(k, nOut_PayCurrency) = InTrades(i, nIn_CCY_PAY)
139                   OutTrades(k, nOut_ReceiveCurrency) = InTrades(i, nIn_CCY_REC)
140                   OutTrades(k, nOut_PayNotional) = -Abs(InTrades(i, nIn_NOMINAL_PAY)) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
141                   OutTrades(k, nOut_ReceiveNotional) = Abs(InTrades(i, nIn_NOMINAL_REC)) * IIf(FlipTrades, -ThisScaleFactor, ThisScaleFactor)
142                   OutTrades(k, nOut_PayBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_PAY_COUPON_DTROLL)))
143                   OutTrades(k, nOut_ReceiveBDC) = ParseAirbusBDC(CStr(InTrades(i, nIn_REC_COUPON_DTROLL)))
144                   PayIsFixed = InTrades(i, nIn_SPREAD_PAY) = "FIX"
145                   ReceiveIsFixed = InTrades(i, nIn_SPREAD_REC) = "FIX"
146                   OutTrades(k, nOut_PayIsFixed) = PayIsFixed
147                   OutTrades(k, nOut_ReceiveIsFixed) = ReceiveIsFixed
148                   If PayIsFixed Then
149                       OutTrades(k, nOut_PayCoupon) = StringWithCommaToNumber(InTrades(i, nIn_INDEX_PAY)) / 100
150                       OutTrades(k, nOut_PayFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PAY_COUPON_FREQ)))
151                       OutTrades(k, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_PAY)), False, True)
152                   Else
153                       OutTrades(k, nOut_PayCoupon) = StringWithCommaToNumber(InTrades(i, nIn_SPREAD_PAY)) / 100
154                       OutTrades(k, nOut_PayFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_PAY_COUPON_FREQ)))
155                       OutTrades(k, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_PAY)), True, True)
156                   End If
157                   If ReceiveIsFixed Then
158                       OutTrades(k, nOut_ReceiveCoupon) = StringWithCommaToNumber(InTrades(i, nIn_INDEX_REC)) / 100
159                       OutTrades(k, nOut_ReceiveFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_REC_COUPON_FREQ)))
160                       OutTrades(k, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_REC)), False, True)
161                   Else
162                       OutTrades(k, nOut_ReceiveCoupon) = StringWithCommaToNumber(InTrades(i, nIn_SPREAD_REC)) / 100
163                       OutTrades(k, nOut_ReceiveFrequency) = ParseAirbusFrequency(CStr(InTrades(i, nIn_REC_COUPON_FREQ)))
164                       OutTrades(k, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, nIn_BASIS_REC)), True, True)
165                   End If
166               Case Else
167                   Throw "Unrecognised ValuationFunction: " + VF
168               End Select
169           End If
170       Next i

171       If k < sNRows(OutTrades) Then OutTrades = sSubArray(OutTrades, 1, 1, k)

          'Now deal with amortising trades...
172       If AnyIRDs Then
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
173           GrabAmortisationData AmortTradeIDs, AmortStartDates, AmortPayRecs, AmortNotionals, AmortMap, twb
174           AmortMapCol1 = sSubArray(AmortMap, 1, 1, , 1)
175           For i = 3 To sNRows(OutTrades)
176               Select Case OutTrades(i, nOut_ValuationFunction)
                  Case "InterestRateSwap", "CrossCurrencySwap"
177                   TradeID = OutTrades(i, nOut_TradeID)
178                   MatchRes = sMatch(CDbl(TradeID), AmortMapCol1, True)
179                   If IsNumber(MatchRes) Then
180                       TheseNotionals = sSubArray(AmortNotionals, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
181                       ThesePayRec = sSubArray(AmortPayRecs, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
182                       TheseReceiveNotionals = sMChoose(TheseNotionals, sArrayEquals(ThesePayRec, "REC"))
183                       ThesePayNotionals = sMChoose(TheseNotionals, sArrayEquals(ThesePayRec, "PAY"))
184                       TheseNotionalDates = sSubArray(AmortStartDates, AmortMap(MatchRes, 2), 1, AmortMap(MatchRes, 3), 1)
185                       TheseReceiveNotionalDates = sMChoose(TheseNotionalDates, sArrayEquals(ThesePayRec, "REC"))
186                       ThesePayNotionalDates = sMChoose(TheseNotionalDates, sArrayEquals(ThesePayRec, "PAY"))

                          Dim FlipAndScale As Double
187                       FlipAndScale = IIf(FlipTrades, -1, 1) * TradesScaleFactor
188                       TheseReceiveNotionals = sArrayMultiply(TheseReceiveNotionals, FlipAndScale)
189                       ThesePayNotionals = sArrayMultiply(ThesePayNotionals, -FlipAndScale)

190                       If OutTrades(i, nOut_ValuationFunction) = "CrossCurrencySwap" Then
191                           ThesePayNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_PayFrequency), CStr(OutTrades(i, nOut_PayBDC)), ThesePayNotionalDates, ThesePayNotionals, PortfolioAgeing)
192                           PayNotionalsString = sConcatenateStrings(ThesePayNotionals, ";")
193                           TheseReceiveNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_ReceiveFrequency), CStr(OutTrades(i, nOut_ReceiveBDC)), TheseReceiveNotionalDates, TheseReceiveNotionals, PortfolioAgeing)
194                           ReceiveNotionalsString = sConcatenateStrings(TheseReceiveNotionals, ";")
195                           OutTrades(i, nOut_PayAmortNotionals) = PayNotionalsString
196                           OutTrades(i, nOut_ReceiveAmortNotionals) = ReceiveNotionalsString
197                           OutTrades(i, nOut_PayNotional) = 0
198                           OutTrades(i, nOut_ReceiveNotional) = 0
199                       Else
                              'InterestRateSwap
200                           OutTrades(i, nOut_Notional) = 0
201                           If OutTrades(i, nOut_ReceiveIsFixed) Then
                                  'trade, not flipped, is one on which Airbus receives fixed
202                               ThesePayNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_FloatingFrequency), CStr(OutTrades(i, nOut_FloatingBDC)), ThesePayNotionalDates, ThesePayNotionals, PortfolioAgeing)
203                               PayNotionalsString = sConcatenateStrings(ThesePayNotionals, ";")
204                               TheseReceiveNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_FixedFrequency), CStr(OutTrades(i, nOut_FixedBDC)), TheseReceiveNotionalDates, TheseReceiveNotionals, PortfolioAgeing)
205                               ReceiveNotionalsString = sConcatenateStrings(TheseReceiveNotionals, ";")
206                               OutTrades(i, nOut_FixedAmortNotionals) = ReceiveNotionalsString
207                               OutTrades(i, nOut_FloatingAmortNotionals) = PayNotionalsString
208                           Else
                                  'Airbus pays fixed
209                               ThesePayNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_FixedFrequency), CStr(OutTrades(i, nOut_FixedBDC)), ThesePayNotionalDates, ThesePayNotionals, PortfolioAgeing)
210                               PayNotionalsString = sConcatenateStrings(ThesePayNotionals, ";")
211                               TheseReceiveNotionals = InferNotionalSchedule(CLng(OutTrades(i, nOut_StartDate)), CDbl(OutTrades(i, nOut_EndDate)), OutTrades(i, nOut_FloatingFrequency), CStr(OutTrades(i, nOut_FloatingBDC)), TheseReceiveNotionalDates, TheseReceiveNotionals, PortfolioAgeing)
212                               ReceiveNotionalsString = sConcatenateStrings(TheseReceiveNotionals, ";")
213                               OutTrades(i, nOut_FixedAmortNotionals) = PayNotionalsString
214                               OutTrades(i, nOut_FloatingAmortNotionals) = ReceiveNotionalsString
215                           End If
216                       End If
217                   End If
218               End Select
219           Next i
220       End If

221       If Compress Then
              Dim numCompressedTrades As Long
222           OutTrades = CompressFxForwards(OutTrades, numCompressedTrades)
223           OutTrades = CompressFxOptions(OutTrades, numCompressedTrades + 1)
224       End If
          'Cache results...
225       PreviousCheckSum = ThisCheckSum
226       PreviousResult = OutTrades
227       PreviousTC = TC
228       GetTradesInRFormat = OutTrades

229       If False Then
230           t5 = sElapsedTime()
231           Debug.Print "GetTradesInRFormat entire method", t5 - t1
232           Debug.Print "ChooseVectorFromFilters", t3 - t2
233           Debug.Print "GetColumnsFromTradesWorkbook", t4 - t3
234           Debug.Print "Both of above"; t4 - t2
235       End If

236       Exit Function
ErrHandler:

237       ErrString = "#GetTradesInRFormat (line " & CStr(Erl) & "): " & Err.Description
238       If TradeID <> "" Then
239           ErrString = ErrString & " (TradeID = " & TradeID & ")!"
240       Else
241           ErrString = ErrString & "!"
242       End If
243       GetTradesInRFormat = ErrString
244       PreviousResult = Empty
245       PreviousCheckSum = ""
246       PreviousNumTrades = 0
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompressFxOptions
' Author    : Philip Swannell
' Date      : 26-Dec-2016
' Purpose   : replace all FxOption trades in a portfolio with FxOptionStrip trades
' -----------------------------------------------------------------------------------------------------------------------
Function CompressFxOptions(TheTrades, FirstIndex As Long)
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
2         Headers = sArrayTranspose(sSubArray(TheTrades, 2, 1, 1))
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

14        NumTrades = sNRows(TheTrades) - 2
15        ChooseVector = sArrayEquals(sSubArray(TheTrades, 1, nIn_ValuationFunction, , 1), "FxOption")
16        NumFxOptions = sArrayCount(ChooseVector)
17        numNotFxOptions = NumTrades - NumFxOptions

18        If NumFxOptions = 0 Then
19            CompressFxOptions = TheTrades
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
30            ExtraCol = sArrayStack("CHAR", "Dates")
31            If numNotFxOptions > 0 Then
32                ExtraCol = sArrayStack(ExtraCol, sReshape("", numNotFxOptions, 1))
33            End If
34            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
35            nIn_Dates = sNCols(ReturnArray)
36        End If
37        If Not IsNumber(nIn_Notionals) Then
38            ExtraCol = sArrayStack("CHAR", "Notionals")
39            If numNotFxOptions > 0 Then
40                ExtraCol = sArrayStack(ExtraCol, sReshape("", numNotFxOptions, 1))
41            End If
42            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
43            nIn_Notionals = sNCols(ReturnArray)
44        End If
45        If Not IsNumber(nIn_Strikes) Then
46            ExtraCol = sArrayStack("CHAR", "Strikes")
47            If numNotFxOptions > 0 Then
48                ExtraCol = sArrayStack(ExtraCol, sReshape("", numNotFxOptions, 1))
49            End If
50            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
51            nIn_Strikes = sNCols(ReturnArray)
52        End If

          Dim FxOptionStripTrades
          Dim thisNull

53        FxOptionStripTrades = sReshape("", numFxOptionStrips, sNCols(ReturnArray))
          'Arrrgh. In R columns of data in a dataframe must be of uniform type, _
           so we have to initialise all the elements of FxOptionStripTrades appropriately.
54        For i = 1 To sNCols(FxOptionStripTrades)
55            Select Case (ReturnArray(1, i))
              Case "CHAR"
56                thisNull = ""
57            Case "BOOL"
58                thisNull = False
59            Case "DOUBLE", "DATESTR", "INT"
60                thisNull = 0
61            Case Else
62                Throw "Unrecognised element in top row of TheTrades:" + CStr(ReturnArray(1, i))
63            End Select
64            For j = 1 To numFxOptionStrips
65                FxOptionStripTrades(j, i) = thisNull
66            Next j
67        Next i

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

89        CompressFxOptions = sArrayStack(ReturnArray, FxOptionStripTrades)

90        Exit Function
ErrHandler:
91        Throw "#CompressFxOptions (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompressFxForwards
' Author    : Philip Swannell
' Date      : 01-Jan-2017
' Purpose   : Compress FxForward trades to FxForwardStrip trades rather than to FixedCashflows
'             trades. Advantage is that Capital calculations based on trade notionals are then
'             possible
' -----------------------------------------------------------------------------------------------------------------------
Function CompressFxForwards(TheTrades, ByRef numCompressedTrades As Long)
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
          Dim TheFxOptions As Variant

1         On Error GoTo ErrHandler
2         Headers = sArrayTranspose(sSubArray(TheTrades, 2, 1, 1))
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

14        NumTrades = sNRows(TheTrades) - 2
15        ChooseVector = sArrayEquals(sSubArray(TheTrades, 1, nIn_ValuationFunction, , 1), "FxForward")
16        NumFxForwards = sArrayCount(ChooseVector)
17        NumNotFxForwards = NumTrades - NumFxForwards

18        If NumFxForwards = 0 Then
19            CompressFxForwards = TheTrades
20            Exit Function
21        End If

22        TheFxOptions = sMChoose(TheTrades, ChooseVector)

23        SortResult = sArrayRange(sSubArray(TheFxOptions, 1, nIn_Counterparty, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_ReceiveCurrency, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_PayCurrency, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_EndDate, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_ReceiveNotional, , 1), _
                                   sSubArray(TheFxOptions, 1, nIn_PayNotional, , 1))

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
30            ExtraCol = sArrayStack("CHAR", "Dates")
31            If NumNotFxForwards > 0 Then
32                ExtraCol = sArrayStack(ExtraCol, sReshape("", NumNotFxForwards, 1))
33            End If
34            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
35            nIn_Dates = sNCols(ReturnArray)
36        End If
37        If Not IsNumber(nIn_ReceiveNotionals) Then
38            ExtraCol = sArrayStack("CHAR", "ReceiveNotionals")
39            If NumNotFxForwards > 0 Then
40                ExtraCol = sArrayStack(ExtraCol, sReshape("", NumNotFxForwards, 1))
41            End If
42            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
43            nIn_ReceiveNotionals = sNCols(ReturnArray)
44        End If
45        If Not IsNumber(nIn_PayNotionals) Then
46            ExtraCol = sArrayStack("CHAR", "PayNotionals")
47            If NumNotFxForwards > 0 Then
48                ExtraCol = sArrayStack(ExtraCol, sReshape("", NumNotFxForwards, 1))
49            End If
50            ReturnArray = sArrayRange(ReturnArray, ExtraCol)
51            nIn_PayNotionals = sNCols(ReturnArray)
52        End If

          Dim FxForwardStripTrades
          Dim thisNull

53        FxForwardStripTrades = sReshape("", NumFxForwardStrips, sNCols(ReturnArray))
          'Arrrgh. In R columns of data in a dataframe must be of uniform type, _
           so we have to initialise all the elements of FxForwardStripTrades appropriately.
54        For i = 1 To sNCols(FxForwardStripTrades)
55            Select Case (ReturnArray(1, i))
              Case "CHAR"
56                thisNull = ""
57            Case "BOOL"
58                thisNull = False
59            Case "DOUBLE", "DATESTR", "INT"
60                thisNull = 0
61            Case Else
62                Throw "Unrecognised element in top row of TheTrades:" + CStr(ReturnArray(1, i))
63            End Select
64            For j = 1 To NumFxForwardStrips
65                FxForwardStripTrades(j, i) = thisNull
66            Next j
67        Next i

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
90        CompressFxForwards = sArrayStack(ReturnArray, FxForwardStripTrades)

91        Exit Function
ErrHandler:
92        Throw "#CompressFxForwards (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ChooseVectorFromFilters
' Author    : Philip Swannell
' Date      : 20-May-2015
' Purpose   : Encapsulate trade selection logic
' -----------------------------------------------------------------------------------------------------------------------
Function ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
                                 PortfolioAgeing As Double, ThrowIfNoTradesFound As Boolean, WithFxTrades As Boolean, _
                                 WithRatesTrades As Boolean, CurrenciesToInclude As String, ByRef TC As TradeCount, _
                                 twb As Workbook, fts As Worksheet, AnchorDate As Date)
          Dim ChooseVector
          Dim ChooseVector2
          Dim Col1
          Dim Col2

1         On Error GoTo ErrHandler
2         If LCase(FilterBy1) = "none" Or LCase(Filter1Value) = "all" Then
3             ChooseVector = sReshape(True, GetColumnFromTradesWorkbook("NumTrades", IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate), 1)
4         Else
5             Col1 = GetColumnFromTradesWorkbook(FilterBy1, IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
6             If FilterWillBeTreatedAsRegExp(Filter1Value) Then
7                 If Not VarType(sIsRegMatch(CStr(Filter1Value), "Foo")) = vbBoolean Then Throw "Invalid regular expression: " + Filter1Value
8                 ChooseVector = sIsRegMatch(CStr(Filter1Value), sArrayMakeText(Col1), False)
9                 ChooseVector = sArrayEquals(True, ChooseVector)
10            Else
11                ChooseVector = sArrayEquals(Col1, Filter1Value)
12            End If
13        End If

14        If LCase(FilterBy2) <> "none" And LCase(Filter2Value) <> "all" Then
15            Col2 = GetColumnFromTradesWorkbook(FilterBy2, IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
16            If FilterWillBeTreatedAsRegExp(Filter2Value) Then
17                If Not VarType(sIsRegMatch(CStr(Filter2Value), "Foo")) = vbBoolean Then Throw "Invalid regular expression: " + Filter2Value
18                ChooseVector2 = sIsRegMatch(CStr(Filter2Value), sArrayMakeText(Col2), False)
19                ChooseVector2 = sArrayEquals(True, ChooseVector2)
20            Else
21                ChooseVector2 = sArrayEquals(Col2, Filter2Value)
22            End If
23            ChooseVector = sArrayAnd(ChooseVector, ChooseVector2)
24        End If

25        ApplyCurrenciesToInclude ChooseVector, IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, CurrenciesToInclude, TC, twb, fts, AnchorDate

26        If ThrowIfNoTradesFound Then
27            If sArrayCount(ChooseVector) = 0 Then Throw "No trades found"    ' for " + SummariseFilters(CStr(FilterBy1), Filter1Value, CStr(FilterBy2), Filter2Value)
28        End If

29        ChooseVectorFromFilters = ChooseVector
30        Exit Function
ErrHandler:
31        Throw "#ChooseVectorFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ApplyCurrenciesToInclude
' Author    : Philip Swannell
' Date      : 27-Sep-2016
' Purpose   : Amends a previously calculated ChooseVector so as to exclude those trades
'             where one or more of the currencies of the trade does not match a substring
'             of CurrenciesToInclude.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ApplyCurrenciesToInclude(ByRef ChooseVector, IncludeFutureTrades As Boolean, _
          PortfolioAgeing As Double, WithFxTrades As Boolean, _
          WithRatesTrades As Boolean, CurrenciesToInclude As String, _
          ByRef TC As TradeCount, _
          twb As Workbook, fts As Worksheet, AnchorDate As Date)
          Dim finalCount As Long
          Dim i As Long
          Dim IncludeThese As Variant
          Dim j As Long
          Dim k As Long
          Dim m As Long
          Dim N As Long
          Dim origCount As Long
          Dim ThisTradeData
          Dim ColName1 As String
          Dim ColName2 As String
          Dim ColName3 As String
          Dim ColName4 As String

1         On Error GoTo ErrHandler

2         If IsTradesWorkbook2022Style(twb) Then
3             ColName1 = "Prim Cur"
4             ColName2 = "Sec Cur"
5             ColName3 = "Pay Ccy"
6             ColName4 = "Rcv Ccy"
7         Else
8             ColName1 = "CCY1"
9             ColName2 = "CCY2"
10            ColName3 = "CCY_REC"
11            ColName4 = "CCY_PAY"
12        End If

13        origCount = sArrayCount(ChooseVector)
14        If LCase(CurrenciesToInclude) = "all" Then
15            TC.NumIncluded = origCount
16            TC.NumExcluded = 0
17            TC.Total = origCount
18            Exit Function
19        End If

20        IncludeThese = sSortedArray(sTokeniseString(CurrenciesToInclude))
21        N = sNRows(ChooseVector)
22        m = sNRows(IncludeThese)

23        For j = IIf(WithFxTrades, 1, 3) To IIf(WithRatesTrades, 4, 2)

24            ThisTradeData = GetColumnFromTradesWorkbook(Choose(j, ColName1, ColName2, ColName3, ColName4), IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
25            For i = 1 To N
26                If ChooseVector(i, 1) Then
27                    If VarType(ThisTradeData(i, 1)) = vbString Then
28                        ChooseVector(i, 1) = False
29                        For k = 1 To m
30                            If ThisTradeData(i, 1) = IncludeThese(k, 1) Then
31                                ChooseVector(i, 1) = True
32                                Exit For
33                            End If
34                        Next k
35                    End If
36                End If
37            Next i
38        Next j
39        finalCount = sArrayCount(ChooseVector)
40        TC.NumExcluded = origCount - finalCount
41        TC.NumIncluded = finalCount
42        TC.Total = origCount

43        Exit Function
ErrHandler:
44        Throw "#ApplyCurrenciesToInclude (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FilterWillBeTreatedAsRegExp
' Author    : Philip Swannell
' Date      : 05-Oct-2016
' Purpose   : This is a kludge...
' -----------------------------------------------------------------------------------------------------------------------
Function FilterWillBeTreatedAsRegExp(Filter As Variant) As Boolean
1         If VarType(Filter) = vbString Then
2             If InStr(Filter, "|") > 0 Or InStr(Filter, "^") > 0 Or InStr(Filter, "$") > 0 Then
3                 FilterWillBeTreatedAsRegExp = True
4             End If
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetColumnFromTradesWorkbook
' Author    : Philip Swannell
' Date      : 28-Apr-2015
' Purpose   : Pass in a column header as they appear in the source workbook. Return is an
'             array matching the contents of that column of the workbook, no header row.
'          Can also pass in HeaderName as "AllHeaders" to get a list of headers.
'          Can also pass in HeaderName as "NumTrades" to get back the number of trades in
'          the source workbook
'          Can also pass "TradeIsFrom" to get a column populated with strings tif_Rates, tif_Fx and tif_Future
'          (meaning the Rates sheet of the trades workbook, the Fx sheet of the trades workbook
'           and the FutureTrades sheet of this workbook) - useful since TradeMorphing must
'           avoid morphing trades from the FutureTrades sheet.
'  1 Aug 2016
'         Modified this method to cope with new structure of trades workbook - one sheet
'         for Fx trades and one sheet for rates trades. Return is still a single column
'         with Fx trades above rates trades above future trades.
' 3 March 2022
'         Modified to cope with either 2017 or 2022 vintage trade workbooks, only difference as far as this method
'         is concerned is the worksheet names.
' -----------------------------------------------------------------------------------------------------------------------
Function GetColumnFromTradesWorkbook(HeaderName As Variant, IncludeFutureTrades As Boolean, PortfolioAgeing As Double, _
          WithFxTrades, WithRatesTrades, twb As Workbook, fts As Worksheet, AnchorDate As Date)

          Dim ChooseVector As Variant
          Dim ColNoFx As Variant
          Dim ColNoRates As Variant
          Dim EntireRangeFxNH As Range
          Dim EntireRangeRatesNH As Range
          Dim FutureTradesSettleDates
          Dim HeaderRowFxTP
          Dim HeaderRowRatesTP As Variant
          Dim lo As ListObject
          Dim MatchID As Variant
          Dim NumFutureTrades As Long
          Dim Result
          Dim ResultFutureTrades
          Dim ResultFx
          Dim ResultRates
          Dim Use2022Format As Boolean
          Dim wsFx As Worksheet
          Dim wsRates As Worksheet

1         On Error GoTo ErrHandler

2         If Not (WithFxTrades Or WithRatesTrades) Then Throw "At least one of WithFxTrades and WithRatesTrades must be True"

3         Use2022Format = IsTradesWorkbook2022Style(twb)

4         If WithFxTrades Then
5             Set wsFx = twb.Worksheets(IIf(Use2022Format, SN_FxTrades2, SN_FxTrades))
6             Select Case wsFx.ListObjects.Count
                  Case 0
                      'Assume data starts at A1
7                     Set EntireRangeFxNH = wsFx.Range("A1").CurrentRegion
8                     Set EntireRangeFxNH = EntireRangeFxNH.Offset(1).Resize(EntireRangeFxNH.Rows.Count - 1)
9                     HeaderRowFxTP = Application.WorksheetFunction.Transpose(EntireRangeFxNH.Rows(0).Value)
10                Case 1
11                    Set lo = wsFx.ListObjects(1)
12                    Set EntireRangeFxNH = lo.DataBodyRange
13                    HeaderRowFxTP = Application.WorksheetFunction.Transpose(EntireRangeFxNH.Rows(0).Value)
14                Case Else
15                    Throw "Cannot find trade data in sheet '" + wsFx.Name + " of workbook '" + twb.Name + "'"
16            End Select
17        End If

18        If WithRatesTrades Then
19            Set wsRates = twb.Worksheets(IIf(Use2022Format, SN_RatesTrades2, SN_RatesTrades))
20            Select Case wsRates.ListObjects.Count
                  Case 0
                      'Assume data starts at A1
21                    Set EntireRangeRatesNH = wsRates.Range("A1").CurrentRegion
22                    Set EntireRangeRatesNH = EntireRangeRatesNH.Offset(1).Resize(EntireRangeRatesNH.Rows.Count - 1)
23                    HeaderRowRatesTP = Application.WorksheetFunction.Transpose(EntireRangeRatesNH.Rows(0).Value)
24                Case 1
25                    Set lo = wsRates.ListObjects(1)
26                    Set EntireRangeRatesNH = lo.DataBodyRange
27                    HeaderRowRatesTP = Application.WorksheetFunction.Transpose(EntireRangeRatesNH.Rows(0).Value)
28                Case Else
29                    Throw "Cannot find trade data in sheet '" + wsRates.Name + " of workbook '" + twb.Name + "'"
30            End Select
31        End If

32        If LCase(HeaderName) = "allheaders" Then
33            If WithRatesTrades And WithFxTrades Then
34                GetColumnFromTradesWorkbook = sRemoveDuplicates(sArrayStack(HeaderRowFxTP, HeaderRowRatesTP), True)
35            ElseIf WithFxTrades Then
36                GetColumnFromTradesWorkbook = sSortedArray(HeaderRowFxTP)
37            ElseIf WithRatesTrades Then
38                GetColumnFromTradesWorkbook = sSortedArray(HeaderRowRatesTP)
39            End If
40            Exit Function
41        ElseIf LCase(HeaderName) = "numtrades" Then
42            GetColumnFromTradesWorkbook = 0
43            If WithFxTrades Then
44                GetColumnFromTradesWorkbook = GetColumnFromTradesWorkbook + EntireRangeFxNH.Rows.Count
45            End If
46            If WithRatesTrades Then
47                GetColumnFromTradesWorkbook = GetColumnFromTradesWorkbook + EntireRangeRatesNH.Rows.Count
48            End If

49            If IncludeFutureTrades Then
50                If IsInCollection(fts.Names, "TheTrades") Then
51                    MatchID = ThrowIfError(sMatch("Settle Date", sArrayTranspose(RangeFromSheet(fts, "Headers").Value)))
                      FutureTradesSettleDates = RangeFromSheet(fts, "TheTrades").Columns(MatchID).Value2
52                    ChooseVector = sArrayLessThanOrEqual(FutureTradesSettleDates, AnchorDate + PortfolioAgeing * 365)
53                    GetColumnFromTradesWorkbook = GetColumnFromTradesWorkbook + sArrayCount(ChooseVector)
54                End If
55            End If
56            Exit Function
57        End If

58        If WithFxTrades Then
59            If HeaderName = "TradeIsFrom" Then
60                ResultFx = sReshape(tif_Fx, EntireRangeFxNH.Rows.Count, 1)
61            Else
62                ColNoFx = sMatch(CVar(HeaderName), HeaderRowFxTP)
63                If IsNumeric(ColNoFx) Then
64                    ResultFx = EntireRangeFxNH.Columns(ColNoFx).Value2
65                Else
66                    ResultFx = sReshape(Empty, EntireRangeFxNH.Rows.Count, 1)
67                End If
68            End If
69        End If

70        If WithRatesTrades Then
71            If HeaderName = "TradeIsFrom" Then
72                ResultRates = sReshape(tif_Rates, EntireRangeRatesNH.Rows.Count, 1)
73            Else
74                ColNoRates = sMatch(CVar(HeaderName), HeaderRowRatesTP)
75                If IsNumeric(ColNoRates) Then
76                    ResultRates = EntireRangeRatesNH.Columns(ColNoRates).Value2
77                Else
78                    ResultRates = sReshape(Empty, EntireRangeRatesNH.Rows.Count, 1)
79                End If
80            End If
81        End If

82        If VarType(ColNoFx) = vbString Then
83            If VarType(ColNoRates) = vbString Then
84                Throw "There is no header '" & HeaderName & "' in the trade data"
85            End If
86        End If

87        If WithFxTrades And WithRatesTrades Then
88            Result = sArrayStack(ResultFx, ResultRates)
89        ElseIf WithFxTrades Then
90            Result = ResultFx
91        Else
92            Result = ResultRates
93        End If

94        If IncludeFutureTrades Then
              'We only take trades from the FutureTrades sheet if they have been traded on or before the _
               date defined by AnchorDate + PortfolioAgeing
95            If IsInCollection(fts.Names, "TheTrades") Then
96                MatchID = ThrowIfError(sMatch("Settle Date", sArrayTranspose(RangeFromSheet(fts, "Headers").Value)))
97                FutureTradesSettleDates = RangeFromSheet(fts, "TheTrades").Columns(MatchID).Value2

98                ChooseVector = sArrayLessThanOrEqual(FutureTradesSettleDates, AnchorDate + PortfolioAgeing * 365)
99                NumFutureTrades = sArrayCount(ChooseVector)
100               If NumFutureTrades > 0 Then
101                   MatchID = sMatch(HeaderName, sArrayTranspose(RangeFromSheet(fts, "Headers").Value))
102                   If IsNumber(MatchID) Then
103                       ResultFutureTrades = sMChoose(RangeFromSheet(fts, "TheTrades").Columns(MatchID).Value2, ChooseVector)
104                   ElseIf HeaderName = "TradeIsFrom" Then
105                       ResultFutureTrades = sReshape(tif_Future, NumFutureTrades, 1)
106                   ElseIf LCase(HeaderName) = "includethisone?" Then
107                       ResultFutureTrades = sReshape(True, NumFutureTrades, 1)
108                   Else
109                       ResultFutureTrades = sReshape(Empty, NumFutureTrades, 1)
110                   End If
111                   Result = sArrayStack(Result, ResultFutureTrades)
112               End If
113           End If
114       End If
115       GetColumnFromTradesWorkbook = Result

116       Exit Function
ErrHandler:
117       Throw "#GetColumnFromTradesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UncompressTrades
' Author    : Philip Swannell
' Date      : 08-Jan-2017
' Purpose   : The inverse of (the composition of) CompressFxForwards and CompressFxOptions
' -----------------------------------------------------------------------------------------------------------------------
Function UncompressTrades(InTrades)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NumInTrades As Long
          Dim NumOutTrades As Long
          Dim w As Long

          Dim cn_Counterparty As Variant
          Dim cn_Currency As Variant
          Dim cn_Dates As Variant
          Dim cn_EndDate As Variant
          Dim cn_isCall As Variant
          Dim cn_Notional As Variant
          Dim cn_Notionals As Variant
          Dim cn_PayCurrency As Variant
          Dim cn_PayNotional As Variant
          Dim cn_PayNotionals As Variant
          Dim cn_ReceiveCurrency As Variant
          Dim cn_ReceiveNotional As Variant
          Dim cn_ReceiveNotionals As Variant
          Dim cn_Strike As Variant
          Dim cn_Strikes As Variant
          Dim cn_TradeID As Variant
          Dim cn_ValuationFunction As Variant
          Dim Headers
          Dim OutTrades() As Variant
1         On Error GoTo ErrHandler
          Dim FormatString As String
          Dim NumCols As Long

2         On Error GoTo ErrHandler

3         Force2DArrayR InTrades

4         NumInTrades = sNRows(InTrades)
5         If NumInTrades <= 2 Then
6             UncompressTrades = InTrades
7             Exit Function
8         Else
9             NumInTrades = NumInTrades - 2
10        End If
11        Headers = sArrayTranspose(sSubArray(InTrades, 2, 1, 1))
12        cn_Dates = sMatch("Dates", Headers): If Not IsNumber(cn_Dates) Then Throw ("Cannot find 'Dates' in row 2 of InTrades")
13        cn_ValuationFunction = sMatch("ValuationFunction", Headers): If Not IsNumber(cn_ValuationFunction) Then Throw ("Cannot find 'ValuationFunction' in row 2 of InTrades")

14        For i = 3 To NumInTrades + 2
15            Select Case (InTrades(i, cn_ValuationFunction))
              Case "FxOptionStrip", "FxForwardStrip"
16                NumOutTrades = NumOutTrades + NumTokens(CStr(InTrades(i, cn_Dates)))
17            Case Else
18                NumOutTrades = NumOutTrades + 1
19            End Select
20        Next i

21        If NumOutTrades = NumInTrades Then
22            UncompressTrades = InTrades
23            Exit Function
24        End If

25        cn_TradeID = sMatch("TradeID", Headers): If Not IsNumber(cn_TradeID) Then Throw ("Cannot find 'TradeID' in row 2 of InTrades")

26        cn_Counterparty = sMatch("Counterparty", Headers): If Not IsNumber(cn_Counterparty) Then Throw ("Cannot find 'Counterparty' in row 2 of InTrades")
27        cn_EndDate = sMatch("EndDate", Headers): If Not IsNumber(cn_EndDate) Then Throw ("Cannot find 'EndDate' in row 2 of InTrades")
28        cn_ReceiveCurrency = sMatch("ReceiveCurrency", Headers): If Not IsNumber(cn_ReceiveCurrency) Then Throw ("Cannot find 'ReceiveCurrency' in row 2 of InTrades")
29        cn_PayCurrency = sMatch("PayCurrency", Headers): If Not IsNumber(cn_PayCurrency) Then Throw ("Cannot find 'PayCurrency' in row 2 of InTrades")
30        cn_ReceiveNotional = sMatch("ReceiveNotional", Headers): If Not IsNumber(cn_ReceiveNotional) Then Throw ("Cannot find 'ReceiveNotional' in row 2 of InTrades")
31        cn_PayNotional = sMatch("PayNotional", Headers): If Not IsNumber(cn_PayNotional) Then Throw ("Cannot find 'PayNotional' in row 2 of InTrades")
32        cn_ReceiveNotionals = sMatch("ReceiveNotionals", Headers): If Not IsNumber(cn_ReceiveNotionals) Then Throw ("Cannot find 'ReceiveNotionals' in row 2 of InTrades")
33        cn_PayNotionals = sMatch("PayNotionals", Headers): If Not IsNumber(cn_PayNotionals) Then Throw ("Cannot find 'PayNotionals' in row 2 of InTrades")
34        cn_Currency = sMatch("Currency", Headers): If Not IsNumber(cn_Currency) Then Throw ("Cannot find 'Currency' in row 2 of InTrades")
35        cn_Notional = sMatch("Notional", Headers): If Not IsNumber(cn_Notional) Then Throw ("Cannot find 'Notional' in row 2 of InTrades")
36        cn_Strike = sMatch("Strike", Headers): If Not IsNumber(cn_Strike) Then Throw ("Cannot find 'Strike' in row 2 of InTrades")
37        cn_isCall = sMatch("isCall", Headers): If Not IsNumber(cn_isCall) Then Throw ("Cannot find 'isCall' in row 2 of InTrades")
38        cn_Notionals = sMatch("Notionals", Headers): If Not IsNumber(cn_Notionals) Then Throw ("Cannot find 'Notionals' in row 2 of InTrades")
39        cn_Strikes = sMatch("Strikes", Headers): If Not IsNumber(cn_Strikes) Then Throw ("Cannot find 'Strikes' in row 2 of InTrades")

40        NumCols = sNCols(InTrades)
41        ReDim OutTrades(1 To NumOutTrades + 2, 1 To NumCols)

          'Copy header rows
42        For i = 1 To 2
43            For j = 1 To NumCols
44                OutTrades(i, j) = InTrades(i, j)
45            Next j
46        Next i

47        k = 3
          Dim TheseDates
          Dim TheseNotionals
          Dim ThesePayNotionals
          Dim TheseReceiveNotionals
          Dim TheseStrikes

48        For i = 3 To NumInTrades + 2
49            Select Case (InTrades(i, cn_ValuationFunction))
              Case "FxForwardStrip"
50                TheseDates = sTokeniseString(CStr(InTrades(i, cn_Dates)), ";")
51                TheseReceiveNotionals = sTokeniseString(CStr(InTrades(i, cn_ReceiveNotionals)), ";")
52                ThesePayNotionals = sTokeniseString(CStr(InTrades(i, cn_PayNotionals)), ";")
53                FormatString = FormatStringFromNumRows(sNRows(TheseDates))
54                For w = 1 To sNRows(TheseDates)
55                    OutTrades(k, cn_TradeID) = CStr(InTrades(i, cn_TradeID)) & "_" & Format(w, FormatString)
56                    OutTrades(k, cn_ValuationFunction) = "FxForward"
57                    OutTrades(k, cn_Counterparty) = InTrades(i, cn_Counterparty)
58                    OutTrades(k, cn_EndDate) = CDbl(CDate(TheseDates(w, 1)))
59                    OutTrades(k, cn_ReceiveCurrency) = InTrades(i, cn_ReceiveCurrency)
60                    OutTrades(k, cn_PayCurrency) = InTrades(i, cn_PayCurrency)
61                    OutTrades(k, cn_ReceiveNotional) = CDbl(TheseReceiveNotionals(w, 1))
62                    OutTrades(k, cn_PayNotional) = CDbl(ThesePayNotionals(w, 1))
63                    k = k + 1
64                Next w
65            Case "FxOptionStrip"
66                TheseDates = sTokeniseString(CStr(InTrades(i, cn_Dates)), ";")
67                TheseStrikes = sTokeniseString(CStr(InTrades(i, cn_Strikes)), ";")
68                TheseNotionals = sTokeniseString(CStr(InTrades(i, cn_Notionals)), ";")
69                FormatString = FormatStringFromNumRows(sNRows(TheseDates))
70                For w = 1 To sNRows(TheseDates)
71                    OutTrades(k, cn_TradeID) = CStr(InTrades(i, cn_TradeID)) & "_" & Format(w, FormatString)
72                    OutTrades(k, cn_ValuationFunction) = "FxOption"
73                    OutTrades(k, cn_Counterparty) = InTrades(i, cn_Counterparty)
74                    OutTrades(k, cn_EndDate) = CDbl(CDate(TheseDates(w, 1)))
75                    OutTrades(k, cn_Currency) = InTrades(i, cn_Currency)
76                    OutTrades(k, cn_Notional) = CDbl(TheseNotionals(w, 1))
77                    OutTrades(k, cn_Strike) = CDbl(TheseStrikes(w, 1))
78                    OutTrades(k, cn_isCall) = InTrades(i, cn_isCall)
79                    k = k + 1
80                Next w
81            Case Else
82                For j = 1 To NumCols
83                    OutTrades(k, j) = InTrades(i, j)
84                Next j
85                k = k + 1
86            End Select
87        Next i
88        UncompressTrades = OutTrades

89        Exit Function
ErrHandler:
90        Throw "#UncompressTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatStringFromNumRows
' Author    : Philip Swannell
' Date      : 08-Jan-2017
' Purpose   : we asign a ..._xxx TradeID to the "unpacked" trades. This mathod returns a format string such that
'             lexicographic ordering is equal to nuermical ordering...
' -----------------------------------------------------------------------------------------------------------------------
Private Function FormatStringFromNumRows(NR As Long)
          Dim l As Long
1         On Error GoTo ErrHandler
2         While 10 ^ l <= NR
3             l = l + 1
4         Wend
5         FormatStringFromNumRows = String(l, "0")
6         Exit Function
ErrHandler:
7         Throw "#FormatStringFromNumRows (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumTokens
' Author    : Philip Swannell
' Date      : 08-Jan-2017
' Purpose   : How many tokens in a semi-colon delimited string. Faster than sNRows(sTokeniseString(...
' -----------------------------------------------------------------------------------------------------------------------
Private Function NumTokens(SCDS As String) As Long
1         On Error GoTo ErrHandler
2         NumTokens = Len(SCDS) - Len(Replace(SCDS, ";", "")) + 1
3         Exit Function
ErrHandler:
4         Throw "#NumTokens (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

