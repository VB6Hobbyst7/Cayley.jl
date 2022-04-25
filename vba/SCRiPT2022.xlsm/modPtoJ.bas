Attribute VB_Name = "modPtoJ"
'---------------------------------------------------------------------------------------
' Module    : modPtoR
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : We pass the trades to Julia in a different "layout" than that displayed on the
'             Portfolio sheet. There are more columns and trade direction is given by the
'             sign of notionals rather than encapsulated in strings such as "BuyCap".
'             To see example trades use Menu > Developer Tools > Show trades formatted for Julia
'             which runs the method ShowTradesForJulia
'             Currently we do not have a method to take trades in the format that Julia accepts
'             and convert them back to the format the Portfolio sheet expects...
'---------------------------------------------------------------------------------------
Option Explicit

Sub TestPortfolioTradesToJuliaTrades()
          Dim NumTrades As Long
          Dim Res

1         On Error GoTo ErrHandler

2         Res = PortfolioTradesToJuliaTrades(getTradesRange(NumTrades).Value2, False, False)

3         g sArrayTranspose(Res)
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestPortfolioTradesToJuliaTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PortfolioTradesToJuliaTrades
' Author    : Philip Swannell
' Date      : 12-Apr-2016
' Purpose   : Takes trade data as held on the Portfolio sheet and creates data that can
'             be passed to Julia. The format passed to Julia is more friendly in the sense that
'             the top row of labels better describes the data, at the cost of having a
'             greater number of columns, to cope with a) different trade types needing
'             different trade attributes and b) R dataframes have a single type per column
'             thus we have to use two different columns for single notionals (numbers) and
'             amortising notionals (semi-colon delimited strings)
'
'             Return has a header row
'             If there are no trades, pass InTrades as Empty
'             This method also validates each trade, which duplicates validation code in
'             method ValidateTrades. This is (currently) necessary because ValidateTrades
'             is not called when calling UpdatePortfolioSheet in AutoCalc mode. The extra
'             validation in this method is proably very cheap n terms of calculation time
'             and "belt and braces" versus ValidateTrades.
'---------------------------------------------------------------------------------------
Function PortfolioTradesToJuliaTrades(InTrades As Variant, ThrowOnError, DatesToStrings As Boolean)

          Dim CopyOfErr As String
          Dim i As Long
          Dim j As Long
          Dim nbTrades As Long
          Dim Numeraire As String
          Dim OutTrades() As Variant
          Dim ThisTradeID As String
          Dim ThisTradeType As String
          
          Const nOut_TradeID As Long = 1
          Const nOut_ValuationFunction As Long = 2
          Const nOut_Counterparty As Long = 3
          Const nOut_StartDate As Long = 4
          Const nOut_EndDate As Long = 5
          Const nOut_ReceiveCurrency As Long = 6
          Const nOut_ReceiveNotional As Long = 7
          Const nOut_ReceiveAmortNotionals As Long = 8
          Const nOut_ReceiveCoupon As Long = 9
          Const nOut_ReceiveIndex As Long = 10
          Const nOut_ReceiveFrequency As Long = 11
          Const nOut_ReceiveDCT As Long = 12
          Const nOut_ReceiveBDC As Long = 13
          Const nOut_PayCurrency As Long = 14
          Const nOut_PayNotional As Long = 15
          Const nOut_PayAmortNotionals As Long = 16
          Const nOut_PayCoupon As Long = 17
          Const nOut_PayIndex As Long = 18
          Const nOut_PayFrequency As Long = 19
          Const nOut_PayDCT As Long = 20
          Const nOut_PayBDC As Long = 21
          Const nOut_Currency As Long = 22
          Const nOut_Notional As Long = 23
          Const nOut_Strike As Long = 24
          Const nOut_FixedFrequency As Long = 25
          Const nOut_FixedDCT As Long = 26
          Const nOut_FixedBDC As Long = 27
          Const nOut_FloatingFrequency As Long = 28
          Const nOut_FloatingDCT As Long = 29
          Const nOut_FloatingBDC As Long = 30
          Const nOut_IsCall As Long = 31
          Const nOut_CashflowDates As Long = 32
          Const nOut_CashflowAmounts As Long = 33
          Const nOut_Dates As Long = 34
          Const nOut_Notionals As Long = 35
          Const nOut_Strikes As Long = 36
          Const nOut_ReceiveNotionals As Long = 37
          Const nOut_PayNotionals As Long = 38
          Const nOut_Coupon As Long = 39
          Const nOut_ReceiveLegType As Long = 40
          Const nOut_PayLegType As Long = 41
          Const nOut_Numeraire As Long = 42
          
          Const nOut_NumCols = 42
          
1         On Error GoTo ErrHandler

2         Numeraire = RangeFromMarketDataBook("Config", "Numeraire")

3         If IsEmpty(InTrades) Then nbTrades = 0 Else nbTrades = sNRows(InTrades)

4         ReDim OutTrades(1 To nbTrades + 1, 1 To nOut_NumCols)

5         OutTrades(1, nOut_Numeraire) = "Numeraire=" & Numeraire 'Shameful bodge way to get the info into the file
6         OutTrades(1, nOut_TradeID) = "TradeID"
7         OutTrades(1, nOut_ValuationFunction) = "ValuationFunction"
8         OutTrades(1, nOut_Counterparty) = "Counterparty"
9         OutTrades(1, nOut_StartDate) = "StartDate"
10        OutTrades(1, nOut_EndDate) = "EndDate"
11        OutTrades(1, nOut_Currency) = "Currency"
12        OutTrades(1, nOut_ReceiveCurrency) = "ReceiveCurrency"
13        OutTrades(1, nOut_PayCurrency) = "PayCurrency"
14        OutTrades(1, nOut_Notional) = "Notional"
15        OutTrades(1, nOut_ReceiveNotional) = "ReceiveNotional"
16        OutTrades(1, nOut_PayNotional) = "PayNotional"
17        OutTrades(1, nOut_ReceiveAmortNotionals) = "ReceiveAmortNotionals"
18        OutTrades(1, nOut_PayAmortNotionals) = "PayAmortNotionals"
19        OutTrades(1, nOut_Coupon) = "Coupon"
20        OutTrades(1, nOut_ReceiveCoupon) = "ReceiveCoupon"
21        OutTrades(1, nOut_PayCoupon) = "PayCoupon"
22        OutTrades(1, nOut_Strike) = "Strike"
23        OutTrades(1, nOut_ReceiveFrequency) = "ReceiveFrequency"
24        OutTrades(1, nOut_PayFrequency) = "PayFrequency"
25        OutTrades(1, nOut_FixedFrequency) = "FixedFrequency"
26        OutTrades(1, nOut_FloatingFrequency) = "FloatingFrequency"
27        OutTrades(1, nOut_IsCall) = "IsCall"
28        OutTrades(1, nOut_CashflowDates) = "CashflowDates"
29        OutTrades(1, nOut_CashflowAmounts) = "CashflowAmounts"
30        OutTrades(1, nOut_ReceiveDCT) = "ReceiveDCT"
31        OutTrades(1, nOut_PayDCT) = "PayDCT"
32        OutTrades(1, nOut_FixedDCT) = "FixedDCT"
33        OutTrades(1, nOut_FloatingDCT) = "FloatingDCT"
34        OutTrades(1, nOut_ReceiveBDC) = "ReceiveBDC"
35        OutTrades(1, nOut_PayBDC) = "PayBDC"
36        OutTrades(1, nOut_FixedBDC) = "FixedBDC"
37        OutTrades(1, nOut_FloatingBDC) = "FloatingBDC"
38        OutTrades(1, nOut_Dates) = "Dates"
39        OutTrades(1, nOut_Notionals) = "Notionals"
40        OutTrades(1, nOut_Strikes) = "Strikes"
41        OutTrades(1, nOut_ReceiveNotionals) = "ReceiveNotionals"
42        OutTrades(1, nOut_PayNotionals) = "PayNotionals"
43        OutTrades(1, nOut_ReceiveLegType) = "ReceiveLegType"
44        OutTrades(1, nOut_PayLegType) = "PayLegType"
45        OutTrades(1, nOut_ReceiveIndex) = "ReceiveIndex"
46        OutTrades(1, nOut_PayIndex) = "PayIndex"

47        For i = 1 To nbTrades
48            ThisTradeID = InTrades(i, gCN_TradeID)
49            ThisTradeType = CStr(InTrades(i, gCN_TradeType))
50            j = i + 1
51            OutTrades(j, nOut_TradeID) = InTrades(i, gCN_TradeID)  'TradeID
52            CopyItem OutTrades, j, nOut_Counterparty, InTrades, i, gCN_Counterparty, vbString, ThisTradeID, "Counterparty"
53            OutTrades(j, nOut_ValuationFunction) = InTrades(i, gCN_TradeType)  'ValuationFunction
54            Select Case ThisTradeType
                  Case "InterestRateSwap", "CrossCurrencySwap"
55                    If ThisTradeType = "InterestRateSwap" Then
56                        If Not SafeEquals(InTrades(i, gCN_Ccy1), InTrades(i, gCN_Ccy2)) Then
57                            Throw "Receive and Pay Currencies must be the same for InterestRateSwap, but they are different for Trade ID " + ThisTradeID
58                        End If
59                    End If
60                    If Not IsValidNotional(InTrades(i, gCN_Notional1)) Then Throw "Receive Notional must be positive or zero or a semi-colon delimitted list of positives or zeros"
61                    If Not IsValidNotional(InTrades(i, gCN_Notional2)) Then Throw "Pay Notional must be positive or zero or a semi-colon delimitted list of positives or zeros"
62                    CopyItem OutTrades, j, nOut_StartDate, InTrades, i, gCN_StartDate, vbDouble, ThisTradeID, "Start Date", DatesToStrings
63                    CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
64                    OutTrades(j, nOut_ReceiveCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
65                    If VarType(InTrades(i, gCN_Notional1)) = vbString Then    'Amortising trade
66                        OutTrades(j, nOut_ReceiveNotional) = 0  'ReceiveNotional
67                        OutTrades(j, nOut_ReceiveAmortNotionals) = InTrades(i, gCN_Notional1)  'ReceiveAmortNotional
68                    Else    ' not amortising trade
69                        CopyItem OutTrades, j, nOut_ReceiveNotional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Receive Notional"
70                    End If
71                    CopyItem OutTrades, j, nOut_ReceiveCoupon, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Receive Coupon"
72                    OutTrades(j, nOut_ReceiveIndex) = ValidateIRLegType(CStr(InTrades(i, gCN_LegType1)), gDoValidation)
73                    OutTrades(j, nOut_ReceiveFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq1)), gDoValidation)     'ReceiveFrequency
74                    OutTrades(j, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT1)), OutTrades(j, nOut_ReceiveIndex) <> "Fixed", gDoValidation)    'ReceiveDCT
75                    OutTrades(j, nOut_ReceiveBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), gDoValidation)    'ReceiveBDC
76                    OutTrades(j, nOut_PayCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), gDoValidation)
77                    If VarType(InTrades(i, gCN_Notional2)) = vbString Then    'Amortising trade
78                        OutTrades(j, nOut_PayNotional) = 0  'PayNotional
79                        OutTrades(j, nOut_PayAmortNotionals) = FlipNotionalSign(CStr(InTrades(i, gCN_Notional2)))    'PayAmortNotional
80                    ElseIf VarType(InTrades(i, gCN_Notional2)) = vbDouble Then  ' not amortising trade
81                        CopyItem OutTrades, j, nOut_PayNotional, InTrades, i, gCN_Notional2, vbDouble, ThisTradeID, "Pay Notional"
82                        OutTrades(j, nOut_PayNotional) = -OutTrades(j, nOut_PayNotional)
83                    Else
84                        Throw "Invalid Pay Notional"
85                    End If
86                    CopyItem OutTrades, j, nOut_PayCoupon, InTrades, i, gCN_Rate2, vbDouble, ThisTradeID, "Pay Coupon"
87                    OutTrades(j, nOut_PayIndex) = ValidateIRLegType(CStr(InTrades(i, gCN_LegType2)), gDoValidation)
88                    OutTrades(j, nOut_PayFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq2)), gDoValidation)    'PayFrequency
89                    OutTrades(j, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT2)), OutTrades(j, nOut_PayIndex) <> "Fixed", gDoValidation) 'PayDCT
90                    OutTrades(j, nOut_PayBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC2)), gDoValidation)    'PayBDC
91                Case "FxForward"
92                    CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
93                    OutTrades(j, nOut_ReceiveCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
94                    OutTrades(j, nOut_PayCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), True)
95                    CopyItem OutTrades, j, nOut_ReceiveNotional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Receive Notional"
96                    If OutTrades(j, nOut_ReceiveNotional) < 0 Then Throw "Receive Notional must be positive or zero"
97                    CopyItem OutTrades, j, nOut_PayNotional, InTrades, i, gCN_Notional2, vbDouble, ThisTradeID, "Pay Notional"
98                    If InTrades(i, gCN_Notional2) < 0 Then Throw "Pay Notional must be positive or zero"
                      'Flip sign on Pay side
99                    OutTrades(j, nOut_PayNotional) = -OutTrades(j, nOut_PayNotional)
100               Case "FxForwardStrip"
101                   ValidateFxStrip ThisTradeID, InTrades(i, gCN_EndDate), InTrades(i, gCN_Notional1), InTrades(i, gCN_Notional2), False
102                   CopyItem OutTrades, j, nOut_Dates, InTrades, i, gCN_EndDate, vbString, ThisTradeID, "End Date"
103                   OutTrades(j, nOut_ReceiveCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
104                   OutTrades(j, nOut_PayCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), True)
105                   CopyItem OutTrades, j, nOut_ReceiveNotionals, InTrades, i, gCN_Notional1, vbString, ThisTradeID, "Receive Notional"
106                   CopyItem OutTrades, j, nOut_PayNotionals, InTrades, i, gCN_Notional2, vbString, ThisTradeID, "Pay Notional"
                      'Flip sign on Pay side
107                   OutTrades(j, nOut_PayNotionals) = FlipSCDS(CStr(OutTrades(j, nOut_PayNotionals)))
108               Case "FxOption"
109                   CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
110                   If VarType(InTrades(i, gCN_Notional2)) <> vbDouble Then Throw "Pay Notional must be a Number"
111                   If InTrades(i, gCN_Notional2) < 0 Then Throw "Pay Notional must be positive or zero"
112                   If VarType(InTrades(i, gCN_Notional1)) <> vbDouble Then Throw "Receive Notional must be a Number"
113                   If InTrades(i, gCN_Notional1) < 0 Then Throw "Receive Notional must be positive or zero"
114                   If CStr(InTrades(i, gCN_Ccy1)) = Numeraire Then
115                       OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), True)
116                       CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional2, vbDouble, ThisTradeID, "Notional"
117                       If InTrades(i, gCN_Notional2) = 0 Then Throw "Pay Notional cannot be zero"
118                       OutTrades(j, nOut_Strike) = InTrades(i, gCN_Notional1) / InTrades(i, gCN_Notional2)  'Strike
119                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "SellCall", "SellPut"
120                               OutTrades(j, nOut_Notional) = -OutTrades(j, nOut_Notional)    'flip sign of Notional for sold option position, have already tested that notionals are numbers
121                           Case "BuyPut", "BuyCall"
                                  'Nothing to do
122                           Case Else
123                               Throw "Unrecognised option style. Allowed values: SellCall, SellPut, BuyCall, BuyPut"
124                       End Select
125                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "BuyPut", "SellPut"
126                               OutTrades(j, nOut_IsCall) = True  'IsCall - The option as booked is a put on the receive currency i.e. put on numeraire i.e. CALL on the non-numeraire Currency
127                           Case Else
128                               OutTrades(j, nOut_IsCall) = False  'IsCall
129                       End Select
130                   ElseIf CStr(InTrades(i, gCN_Ccy2)) = Numeraire Then
131                       OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)

132                       If InTrades(i, gCN_Notional1) = 0 Then Throw "Receive Notional cannot be zero"
133                       OutTrades(j, nOut_Strike) = InTrades(i, gCN_Notional2) / InTrades(i, gCN_Notional1)  'Strike. Have already tested that both notionals are numeric and that the denominator is non zero.
134                       OutTrades(j, nOut_Notional) = InTrades(i, gCN_Notional1)  'Notional
135                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "SellCall", "SellPut"
136                               OutTrades(j, nOut_Notional) = -OutTrades(j, nOut_Notional)    'flip sign of Notional for sold option position
137                           Case "BuyCall", "BuyPut"
138                           Case Else
139                               Throw "Unrecognised option style. Allowed values: SellCall, SellPut, BuyCall, BuyPut"
140                       End Select
141                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "BuyCall", "SellCall"
142                               OutTrades(j, nOut_IsCall) = True  'IsCall - The option as booked is a call on the receive currency i.e. call on non-numeraire so we book as call
143                           Case Else
144                               OutTrades(j, nOut_IsCall) = False  'IsCall
145                       End Select
146                   Else
147                       Throw "FxOptions on cross rates not yet supported i.e. one of the two currencies must be the numeraire currency (" + Numeraire + "), TradeID = " + CStr(InTrades(i, gCN_TradeID))
148                   End If
149               Case "FxOptionStrip"    'Very similar to FxOption but process Dates, Notionals and Strikes as arrays encoded as semi-colon delimited strings
150                   ValidateFxStrip ThisTradeID, InTrades(i, gCN_EndDate), InTrades(i, gCN_Notional1), InTrades(i, gCN_Notional2), True
151                   CopyItem OutTrades, j, nOut_Dates, InTrades, i, gCN_EndDate, vbString, ThisTradeID, "End Date"
152                   If CStr(InTrades(i, gCN_Ccy1)) = Numeraire Then
153                       OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), True)
154                       CopyItem OutTrades, j, nOut_Notionals, InTrades, i, gCN_Notional2, vbString, ThisTradeID, "Notional 1"
155                       OutTrades(j, nOut_Strikes) = DivideSCDS(CStr(InTrades(i, gCN_Notional1)), CStr(InTrades(i, gCN_Notional2)))
156                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "SellCall", "SellPut"
157                               OutTrades(j, nOut_Notionals) = FlipSCDS(CStr(OutTrades(j, nOut_Notionals)))    'flip sign of Notional for sold option position.
158                           Case "BuyPut", "BuyCall"
                                  'Nothing to do
159                           Case Else
160                               Throw "Unrecognised option style. Allowed values: SellCall, SellPut, BuyCall, BuyPut"
161                       End Select
162                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "BuyPut", "SellPut"
163                               OutTrades(j, nOut_IsCall) = True  'IsCall - The option as booked is a put on the receive currency i.e. put on numeraire i.e. CALL on the non-numeraire Currency
164                           Case Else
165                               OutTrades(j, nOut_IsCall) = False  'IsCall
166                       End Select
167                   ElseIf CStr(InTrades(i, gCN_Ccy2)) = Numeraire Then
168                       OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
169                       OutTrades(j, nOut_Strikes) = DivideSCDS(CStr(InTrades(i, gCN_Notional2)), CStr(InTrades(i, gCN_Notional1)))
170                       OutTrades(j, nOut_Notionals) = InTrades(i, gCN_Notional1)
171                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "SellCall", "SellPut"
172                               OutTrades(j, nOut_Notionals) = FlipSCDS(CStr(OutTrades(j, nOut_Notionals)))     'flip sign of Notional for sold option position
173                           Case "BuyCall", "BuyPut"
174                           Case Else
175                               Throw "Unrecognised option style. Allowed values: SellCall, SellPut, BuyCall, BuyPut"
176                       End Select
177                       Select Case CStr(InTrades(i, gCN_LegType1))
                              Case "BuyCall", "SellCall"
178                               OutTrades(j, nOut_IsCall) = True  'IsCall - The option as booked is a call on the receive currency i.e. call on non-numeraire so we book as call
179                           Case Else
180                               OutTrades(j, nOut_IsCall) = False  'IsCall
181                       End Select
182                   Else
183                       Throw "FxOptions on cross rates not yet supported i.e. one of the two currencies must be the numeraire currency (" + Numeraire + "), TradeID = " + CStr(InTrades(i, gCN_TradeID))
184                   End If
185               Case "Swaption"
186                   CopyItem OutTrades, j, nOut_StartDate, InTrades, i, gCN_StartDate, vbDouble, ThisTradeID, "Start Date", DatesToStrings
187                   CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
188                   OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
189                   If VarType(InTrades(i, gCN_Notional1)) <> vbDouble Then Throw "Notional must be a Number"
190                   If InTrades(i, gCN_Notional1) < 0 Then Throw "Notional must be positive or zero"
191                   Select Case CStr(InTrades(i, gCN_LegType1))
                          Case "BuyReceivers"
192                           OutTrades(j, nOut_Notional) = InTrades(i, gCN_Notional1)  'Notional, have already tested for is number
193                           OutTrades(j, nOut_IsCall) = False  'IsCall
194                       Case "SellReceivers"
195                           OutTrades(j, nOut_Notional) = InTrades(i, gCN_Notional1) * -1  'Notional, have already tested for is number
196                           OutTrades(j, nOut_IsCall) = False  'IsCall
197                       Case "BuyPayers"
198                           OutTrades(j, nOut_Notional) = InTrades(i, gCN_Notional1)  'Notional, have already tested for is number
199                           OutTrades(j, nOut_IsCall) = True  'IsCall
200                       Case "SellPayers"
201                           OutTrades(j, nOut_Notional) = InTrades(i, gCN_Notional1) * -1  'Notional, have already tested for is number
202                           OutTrades(j, nOut_IsCall) = True  'IsCall
203                       Case Else
204                           Throw "Unrecognised option style. Allowed values: BuyReceivers, SellReceivers, BuyPayers, SellPayers"
205                   End Select
206                   CopyItem OutTrades, j, nOut_Strike, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Strike"
207                   OutTrades(j, nOut_FixedFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq1)), gDoValidation)    'FixedFrequency
208                   OutTrades(j, nOut_FloatingFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq2)), gDoValidation)    'FloatingFrequency
209                   OutTrades(j, nOut_FixedDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT1)), False, gDoValidation)   'FixedDCT
210                   OutTrades(j, nOut_FloatingDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT2)), True, gDoValidation)   'FloatingDCT
211                   OutTrades(j, nOut_FixedBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), gDoValidation)    'FixedBDC
212                   OutTrades(j, nOut_FloatingBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC2)), gDoValidation)    'FloatingBDC
213               Case "CapFloor"
214                   CopyItem OutTrades, j, nOut_StartDate, InTrades, i, gCN_StartDate, vbDouble, ThisTradeID, "Start Date", DatesToStrings
215                   CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
216                   OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), True)
217                   Select Case CStr(InTrades(i, gCN_LegType1))
                          Case "BuyCap"
218                           CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional"
219                           OutTrades(j, nOut_IsCall) = True  'IsCall
220                       Case "SellCap"
221                           CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional"
222                           OutTrades(j, nOut_Notional) = -OutTrades(j, nOut_Notional)  'Flip sign
223                           OutTrades(j, nOut_IsCall) = True  'IsCall
224                       Case "BuyFloor"
225                           CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional"
226                           OutTrades(j, nOut_IsCall) = False  'IsCall
227                       Case "SellFloor"
228                           CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional"
229                           OutTrades(j, nOut_Notional) = -OutTrades(j, nOut_Notional)  'Flip sign
230                           OutTrades(j, nOut_IsCall) = False  'IsCall
231                       Case Else
232                           Throw "Unrecognised option style for trade " + ThisTradeID
233                   End Select
234                   If InTrades(i, gCN_Notional1) < 0 Then Throw "Notional must be positive or zero"
235                   CopyItem OutTrades, j, nOut_Strike, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Strike"
236                   OutTrades(j, nOut_FloatingFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq1)), gDoValidation)    'FloatingFrequency
237                   OutTrades(j, nOut_FloatingDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT1)), True, gDoValidation)   'FloatingDCT
238                   OutTrades(j, nOut_FloatingBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), gDoValidation)    'FloatingBDC
239               Case "FixedCashflows"
240                   ValidateFixedCashflows ThisTradeID, InTrades(i, gCN_EndDate), InTrades(i, gCN_Notional1)
241                   OutTrades(j, nOut_CashflowDates) = InTrades(i, gCN_EndDate)  'CashflowDates
242                   OutTrades(j, nOut_CashflowAmounts) = InTrades(i, gCN_Notional1)  'CashflowAmounts
243                   OutTrades(j, nOut_Currency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), gDoValidation)
244               Case "InflationZCSwap"
245                   CopyItem OutTrades, j, nOut_StartDate, InTrades, i, gCN_StartDate, vbDouble, ThisTradeID, "Start Date", DatesToStrings
246                   CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
247                   OutTrades(j, nOut_Currency) = ValidateInflationIndex(CStr(InTrades(i, gCN_Ccy1)), True)    'Inflation Index
248                   CopyItem OutTrades, j, nOut_Notional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional"
249                   If Not SafeEquals(InTrades(i, gCN_Notional2), InTrades(i, gCN_Notional1)) Then
250                       Throw "Notionals must be the same for both legs, but they are different for Trade ID " + ThisTradeID
251                   ElseIf InTrades(i, gCN_Notional1) < 0 Then
252                       Throw "Notional must be greater than or equal to zero, but it's not for Trade ID " + ThisTradeID
253                   End If
254                   Select Case CStr(InTrades(i, gCN_LegType1)) & "|" & CStr(InTrades(i, gCN_LegType2))
                          Case "Fixed|Index"
255                           CopyItem OutTrades, j, nOut_Coupon, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Rate 1"
256                       Case "Index|Fixed"
257                           CopyItem OutTrades, j, nOut_Coupon, InTrades, i, gCN_Rate2, vbDouble, ThisTradeID, "Rate 2"
258                           OutTrades(j, nOut_Notional) = -OutTrades(j, nOut_Notional)
259                       Case Else
260                           Throw "Invalid values for 'Is Fixed? 1' and 'Is Fixed 2' one must read 'Index' and the other 'Fixed' but that's not the case for Trade ID " + ThisTradeID
261                   End Select
262                   If InTrades(i, gCN_LegType1) = "Fixed" Then
263                       OutTrades(j, nOut_FixedBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), True)  'FixedBDC
264                       OutTrades(j, nOut_FloatingBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC2)), True)    'FloatingBDC
265                   Else
266                       OutTrades(j, nOut_FixedBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC2)), True)    'FloatingBDC
267                       OutTrades(j, nOut_FloatingBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), True)  'FixedBDC
268                   End If
269               Case "InflationYoYSwap"
270                   CopyItem OutTrades, j, nOut_StartDate, InTrades, i, gCN_StartDate, vbDouble, ThisTradeID, "Start Date", DatesToStrings
271                   CopyItem OutTrades, j, nOut_EndDate, InTrades, i, gCN_EndDate, vbDouble, ThisTradeID, "Maturity Date", DatesToStrings
272                   CopyItem OutTrades, j, nOut_ReceiveNotional, InTrades, i, gCN_Notional1, vbDouble, ThisTradeID, "Notional 1"
273                   If Not SafeEquals(InTrades(i, gCN_Notional2), InTrades(i, gCN_Notional1)) Then
274                       Throw "Notionals must be the same for both legs, but they are different for Trade ID " + ThisTradeID
275                   ElseIf InTrades(i, gCN_Notional1) < 0 Then
276                       Throw "Notional must be greater than or equal to zero, but it's not for Trade ID " + ThisTradeID
277                   End If
278                   Select Case CStr(InTrades(i, gCN_LegType1)) & "|" & CStr(InTrades(i, gCN_LegType2))
                          Case "Index|Fixed", "Index|Floating", "Fixed|Index", "Floating|Index"
279                       Case Else
280                           Throw "Invalid values for 'Is Fixed? 1' and 'Is Fixed 2' one must read 'Index' and the other either 'Fixed' or 'Floating' but that's not the case for Trade ID " + ThisTradeID
281                   End Select
282                   Select Case CStr(InTrades(i, gCN_LegType1))
                          Case "Fixed"
283                           OutTrades(j, nOut_ReceiveLegType) = "FixedLeg"
284                           CopyItem OutTrades, j, nOut_ReceiveCoupon, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Rate 1"
285                           OutTrades(j, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT1)), False, gDoValidation)  'ReceiveDCT
286                           OutTrades(j, nOut_ReceiveCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), gDoValidation)
287                       Case "Floating"
288                           OutTrades(j, nOut_ReceiveLegType) = "FloatingLeg"
289                           CopyItem OutTrades, j, nOut_ReceiveCoupon, InTrades, i, gCN_Rate1, vbDouble, ThisTradeID, "Rate 1"
290                           OutTrades(j, nOut_ReceiveDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT1)), True, gDoValidation)  'ReceiveDCT
291                           OutTrades(j, nOut_ReceiveCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy1)), gDoValidation)
292                       Case "Index"
293                           OutTrades(j, nOut_ReceiveLegType) = "InflationYoYLeg"
294                           If CStr(InTrades(i, gCN_DCT1)) <> "ActB/ActB" Then Throw "DCT 1 must be 'ActB/ActB' on the Index leg of InflationYoYSwap. TradeID = " & ThisTradeID
295                           OutTrades(j, nOut_ReceiveDCT) = InTrades(i, gCN_DCT1)    'ReceiveDCT
                              'Don't copy coupon across as we don't allow a margin on the index side
296                           OutTrades(j, nOut_ReceiveCurrency) = ValidateInflationIndex(CStr(InTrades(i, gCN_Ccy1)), gDoValidation)
297                       Case Else
298                           Throw "Invalid value for 'Is Fixed? 1' for Trade ID " + ThisTradeID
299                   End Select
300                   OutTrades(j, nOut_ReceiveFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq1)), gDoValidation)    'ReceiveFrequency
301                   OutTrades(j, nOut_ReceiveBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC1)), gDoValidation)     'ReceiveBDC
302                   CopyItem OutTrades, j, nOut_PayNotional, InTrades, i, gCN_Notional2, vbDouble, ThisTradeID, "Notional 2"
303                   OutTrades(j, nOut_PayNotional) = -OutTrades(j, nOut_PayNotional)    'Flip sign on pay leg
304                   Select Case CStr(InTrades(i, gCN_LegType2))
                          Case "Fixed"
305                           OutTrades(j, nOut_PayLegType) = "FixedLeg"
306                           CopyItem OutTrades, j, nOut_PayCoupon, InTrades, i, gCN_Rate2, vbDouble, ThisTradeID, "Rate 2"
307                           OutTrades(j, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT2)), False, gDoValidation)  'PayDCT
308                           OutTrades(j, nOut_PayCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), gDoValidation)
309                       Case "Floating"
310                           OutTrades(j, nOut_PayLegType) = "FloatingLeg"
311                           CopyItem OutTrades, j, nOut_PayCoupon, InTrades, i, gCN_Rate2, vbDouble, ThisTradeID, "Rate 2"
312                           OutTrades(j, nOut_PayDCT) = sParseDCT(CStr(InTrades(i, gCN_DCT2)), True, gDoValidation)  'PayDCT
313                           OutTrades(j, nOut_PayCurrency) = ValidateCurrency(CStr(InTrades(i, gCN_Ccy2)), gDoValidation)
314                       Case "Index"
315                           OutTrades(j, nOut_PayLegType) = "InflationYoYLeg"
316                           If CStr(InTrades(i, gCN_DCT2)) <> "ActB/ActB" Then Throw "DCT 1 must be 'ActB/ActB' on the Index leg of InflationYoYSwap. TradeID = " & ThisTradeID
317                           OutTrades(j, nOut_PayDCT) = InTrades(i, gCN_DCT2)
                              'Don't copy coupon across as we don't allow a margin on the index side
318                           OutTrades(j, nOut_PayCurrency) = ValidateInflationIndex(CStr(InTrades(i, gCN_Ccy2)), True)
319                       Case Else
320                           Throw "Invalid value for 'Is Fixed? 1' for Trade ID " + ThisTradeID
321                   End Select
322                   OutTrades(j, nOut_PayFrequency) = sParseFrequencyString(CStr(InTrades(i, gCN_Freq2)), gDoValidation)    'PayFrequency
323                   OutTrades(j, nOut_PayBDC) = ValidateBDC(CStr(InTrades(i, gCN_BDC2)), gDoValidation)    'PayBDC
324               Case Else
325                   Throw "Unrecognised value in TradeType column: " + ThisTradeType
326           End Select
327       Next i

328       PortfolioTradesToJuliaTrades = OutTrades

329       Exit Function
ErrHandler:
330       CopyOfErr = Err.Description
331       If InStr(CopyOfErr, ThisTradeID) = 0 Then
              'replacement of characters in line below ensures that annotation of the error is respected in method SomethingWentWrong
332           CopyOfErr = "#PortfolioTradesToJuliaTrades (line " & CStr(Erl) + "): " + "(TradeID = " + CStr(ThisTradeID) + ") " & Replace(Replace(Replace(CopyOfErr, "#", ""), "!", ""), ":", "") & "!"
333       Else
334           CopyOfErr = "#PortfolioTradesToJuliaTrades (line " & CStr(Erl) + "): " & CopyOfErr & "!"
335       End If
336       If ThrowOnError Then
337           Throw CopyOfErr
338       Else
339           PortfolioTradesToJuliaTrades = CopyOfErr
340       End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : DivideSCDS
' Author    : Philip Swannell
' Date      : 25-Dec-2016
' Purpose   : Divide semi-colon-delimited strings eg DivideSCDS("1;1;1","1;2;4") = "1;0.5;0.25"
'---------------------------------------------------------------------------------------
Private Function DivideSCDS(String1 As String, String2 As String)
          Dim Array1 As Variant
          Dim Array2 As Variant
          Dim Array3 As Variant
          Dim i As Long

1         On Error GoTo ErrHandler
2         Array1 = sTokeniseString(String1, ";")
3         Array2 = sTokeniseString(String2, ";")
4         Array3 = Array1
5         For i = 1 To sNRows(Array1)
6             If Array2(i, 1) = "0" Then Throw "Divide by zero"
7             Array3(i, 1) = CStr(CDbl(Array1(i, 1)) / CDbl(Array2(i, 1)))
8         Next i
9         DivideSCDS = sConcatenateStrings(Array3, ";")

10        Exit Function
ErrHandler:
11        Throw "#DivideSCDS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MultiplySCDS
' Author    : Philip Swannell
' Date      : 25-Dec-2016
' Purpose   : Multiply semi-colon-delimited strings eg MultiplySCDS("1;2;3","1;2;4") = "1;4;12"
'---------------------------------------------------------------------------------------
Function MultiplySCDS(String1 As String, String2 As String)
          Dim Array1 As Variant
          Dim Array2 As Variant
          Dim Array3 As Variant
          Dim i As Long

1         On Error GoTo ErrHandler
2         Array1 = sTokeniseString(String1, ";")
3         Array2 = sTokeniseString(String2, ";")
4         Array3 = Array1
5         For i = 1 To sNRows(Array1)
6             Array3(i, 1) = CStr(CDbl(Array1(i, 1)) * CDbl(Array2(i, 1)))
7         Next i
8         MultiplySCDS = sConcatenateStrings(Array3, ";")

9         Exit Function
ErrHandler:
10        Throw "#MultiplySCDS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FlipSCDS
' Author    : Philip Swannell
' Date      : 25-Dec-2016
' Purpose   : Change sign of semi-colon delimited string e.g "1;-10" --> "-1,10"
'---------------------------------------------------------------------------------------
Function FlipSCDS(SCDS As String)
          Dim Array1
          Dim i As Long

1         On Error GoTo ErrHandler
2         Array1 = sTokeniseString(SCDS, ";")
3         For i = 1 To sNRows(Array1)
4             Array1(i, 1) = CStr(-CDbl(Array1(i, 1)))
5         Next i
6         FlipSCDS = sConcatenateStrings(Array1, ";")

7         Exit Function
ErrHandler:
8         Throw "#FlipSCDS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateBDC
' Author    : Philip Swannell
' Date      : 10-May-2016
' Purpose   : Validate a day count type
'             CHANGING THIS FUNCTION? Then also make equivalent change to method SupportedBDCs
'---------------------------------------------------------------------------------------
Function ValidateBDC(BDC As String, ThrowOnError As Boolean)
          Dim ErrString As String
1         On Error GoTo ErrHandler
2         If Not gDoValidation Then
3             ValidateBDC = BDC
4             Exit Function
5         End If

6         Select Case BDC
              Case "Mod Foll", "Foll", "Mod Prec", "Prec", "None"
7                 ValidateBDC = BDC
8             Case Else
9                 ErrString = "Invalid business day convention: '" + BDC + "' Valid types are " + sConcatenateStrings(SupportedBDCs())
10                If ThrowOnError Then
11                    Throw ErrString
12                Else
13                    ValidateBDC = "#" + ErrString + "!"
14                End If
15        End Select
16        Exit Function
ErrHandler:
17        Throw "#ValidateBDC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateIRLegType
' Author     : Philip Swannell
' Date       : 16-Nov-2020
' Purpose   : Validate a day count type
'             CHANGING THIS FUNCTION? Then also make equivalent change to method SupportedIRLegTypes
' -----------------------------------------------------------------------------------------------------------------------
Function ValidateIRLegType(IRLegType As String, ThrowOnError As Boolean)
          Dim ErrString As String
1         On Error GoTo ErrHandler
2         If Not gDoValidation Then
3             ValidateIRLegType = IRLegType
4             Exit Function
5         End If

6         Select Case IRLegType
              Case "Fixed", "IBOR", "RFR"
7                 ValidateIRLegType = IRLegType
8             Case Else
9                 ErrString = "Invalid interest rate leg type type: '" + IRLegType + "' Valid types are " + sConcatenateStrings(SupportedIRLegTypes())
10                If ThrowOnError Then
11                    Throw ErrString
12                Else
13                    ValidateIRLegType = "#" + ErrString + "!"
14                End If
15        End Select
16        Exit Function
ErrHandler:
17        Throw "#ValidateIRLegType (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateCurrency
' Author     : Philip Swannell
' Date       : 31-Jan-2019
' Purpose    : Validates a currency string against all the currencies returned by sCurrencies, uses a collection for speed
'              (about 100 times faster than using sMatch) though hard-wired list inside select case might be faster (by factor of2?)
' -----------------------------------------------------------------------------------------------------------------------
Function ValidateCurrency(Ccy As String, ThrowOnError As Boolean)
          Static c As Collection, ErrString As String

1         On Error GoTo ErrHandler
          'Faster to use a case statement for the most common currencies?
2         Select Case Ccy
              Case "USD", "EUR", "JPY", "GBP", "AUD", "CAD", "CHF", "CNH", "SEK"
3                 ValidateCurrency = Ccy
4                 Exit Function
5         End Select

6         If c Is Nothing Then
              Dim AllCcys As Variant
              Dim i As Long
7             AllCcys = sCurrencies(False, False)
8             Set c = New Collection
9             For i = 1 To sNRows(AllCcys)
10                c.Add AllCcys(i, 1), AllCcys(i, 1)
11            Next i
12        End If
13        If IsInCollection(c, Ccy) Then
14            ValidateCurrency = Ccy
15        Else
16            ErrString = "Invalid currency: '" + Ccy + "'"
17            If ThrowOnError Then
18                Throw ErrString
19            Else
20                ValidateCurrency = "#" + ErrString + "!"
21            End If
22        End If

23        Exit Function
ErrHandler:
24        Throw "#ValidateCurrency (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateInflationIndex
' Author    : Philip Swannell
' Date      : 12-May-2017
' Purpose   :
'---------------------------------------------------------------------------------------
Function ValidateInflationIndex(Index As String, ThrowOnError As Boolean)
          Dim Allowed
          Dim ErrString
          Static c As Collection, i As Long

1         On Error GoTo ErrHandler
2         If Not gDoValidation Then
3             ValidateInflationIndex = Index
4             Exit Function
5         End If

6         If c Is Nothing Then
7             Set c = New Collection
8             Allowed = SupportedInflationIndices()
9             For i = 1 To sNRows(Allowed)
10                c.Add "", Allowed(i, 1)
11            Next i
12        End If

13        If IsInCollection(c, Index) Then
14            ValidateInflationIndex = Index
15        Else
16            ErrString = "Invalid inflation index: '" + Index + "'"
17            If ThrowOnError Then
18                Throw ErrString
19            Else
20                ValidateInflationIndex = "#" + ErrString + "!"
21            End If
22        End If
23        Exit Function
ErrHandler:
24        Throw "#ValidateInflationIndex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SafeEquals(a, b)
1         On Error GoTo ErrHandler
2         SafeEquals = a = b
3         Exit Function
ErrHandler:
4         SafeEquals = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : CopyItem
' Author    : Philip Swannell
' Date      : 15-Apr-2016
' Purpose   : Copy items from an address in the array InTrades to an address in the array
'             OutTrades and also do type checking.
'---------------------------------------------------------------------------------------
Function CopyItem(ByRef OutTrades, WriteRow, WriteColumn, ByRef InTrades, ReadRow, ReadColumn, TheVarType, TradeID, TradeAttributeName, Optional DatesToStrings As Boolean)
1         If (VarType(InTrades(ReadRow, ReadColumn)) <> TheVarType) And gDoValidation Then
2             Throw "Invalid " + TradeAttributeName + " for trade " + TradeID + ": " + CStr2(InTrades(ReadRow, ReadColumn))
3         Else
4             If DatesToStrings Then
5                 OutTrades(WriteRow, WriteColumn) = FormatAsDate(InTrades(ReadRow, ReadColumn), "YYYY-MM-DD")
6             Else
7                 OutTrades(WriteRow, WriteColumn) = InTrades(ReadRow, ReadColumn)
8             End If
9         End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CStr2
' Author    : Philip Swannell
' Date      : 12-May-2017
' Purpose   : Cast to string with special handling for Empty and nullstring, for more friendly error messages
'---------------------------------------------------------------------------------------
Private Function CStr2(x) As String
1         On Error GoTo ErrHandler
2         If IsEmpty(x) Then
3             CStr2 = "Empty Cell"
4         Else
5             CStr2 = CStr(x)
6             If CStr2 = "" Then
7                 CStr2 = "Zero-length string"
8             Else
9                 CStr2 = "'" + CStr2 + "'"
10            End If
11        End If
12        Exit Function
ErrHandler:
13        Throw "#CStr2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShowTradesForJulia
' Author    : Philip Swannell
' Date      : 14-Apr-2016
' Purpose   : Display the trades in the format in which we pass them to Julia, mainly for debugging.
'---------------------------------------------------------------------------------------
Sub ShowTradesForJulia(Optional AsRows As Boolean = False)
          Dim Headers
          Dim JuliaTrades As Variant
          Dim NumTrades As Long
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TR As Range
          Dim TR2 As Range

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set TR = getTradesRange(NumTrades)
4         If NumTrades = 0 Then Exit Sub
5         JuliaTrades = PortfolioTradesToJuliaTrades(TR.Value2, True, False)
6         If AsRows Then
7             ExamineInSheet JuliaTrades
8             Set TR2 = ActiveSheet.UsedRange
9             Headers = sArrayTranspose(TR2.Rows(1).Value)
10            TR2.Columns(sMatch("StartDate", Headers)).NumberFormat = NF_Date
11            TR2.Columns(sMatch("EndDate", Headers)).NumberFormat = NF_Date
12            TR2.Columns(sMatch("Notional", Headers)).NumberFormat = NF_Comma0dp
13            TR2.Columns(sMatch("PayNotional", Headers)).NumberFormat = NF_Comma0dp
14            TR2.Columns(sMatch("ReceiveNotional", Headers)).NumberFormat = NF_Comma0dp
15        Else
16            JuliaTrades = sArrayTranspose(JuliaTrades)
17            ExamineInSheet JuliaTrades
18            Set TR2 = ActiveSheet.UsedRange
19            Headers = TR2.Columns(1).Value
20            TR2.Rows(sMatch("StartDate", Headers)).NumberFormat = NF_Date
21            TR2.Rows(sMatch("EndDate", Headers)).NumberFormat = NF_Date
22            TR2.Rows(sMatch("Notional", Headers)).NumberFormat = NF_Comma0dp
23            TR2.Rows(sMatch("PayNotional", Headers)).NumberFormat = NF_Comma0dp
24            TR2.Rows(sMatch("ReceiveNotional", Headers)).NumberFormat = NF_Comma0dp
25        End If
26        AutoFitColumns TR2, , , 30
27        TR2.HorizontalAlignment = xlHAlignLeft
28        TR2.Parent.Parent.Save

29        Exit Sub
ErrHandler:
30        SomethingWentWrong "#ShowTradesForJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
