Attribute VB_Name = "modJtoP"
'---------------------------------------------------------------------------------------
' Module    : modRtoP
' Author    : Philip Swannell
' Date      : 23-Dec-2016
' Purpose   : Code to translate trades in the form that we send them to the Julia code to trades
'             in the format that's displayed on the Portfolio sheet of this workbook
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : SafeMatch
' Author    : Hermione Glyn
' Date      : 10-Aug-2016
' Purpose   : Wrapper to sMatch but when not found return 0.
'---------------------------------------------------------------------------------------
Private Function SafeMatch(HeadersT As Variant, Header As String)
          Dim matchRes

1         On Error GoTo ErrHandler

2         matchRes = sMatch(Header, HeadersT)
3         If IsNumber(matchRes) Then
4             SafeMatch = matchRes
5         Else
6             SafeMatch = 0    'we don't throw at this stage since we don't yet know if we will want to read from this column
7             Exit Function
8         End If

9         Exit Function
ErrHandler:
10        Throw "#SafeMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
'---------------------------------------------------------------------------------------
' Procedure : sParseFrequencyNumber
' Author    : Hermione Glyn
' Date      : 10-Aug-2016
' Purpose   : Convert a numerical description of payment frequency to a string
'             (reverse of sParseFrequencyString)
'---------------------------------------------------------------------------------------
Function sParseFrequencyNumber(FrequencyNumber As Long, ThrowOnError As Boolean) As String
1         On Error GoTo ErrHandler
          Const ErrString = "Frequency not recognised. Allowed values: 1, 2, 4, 12."

2         Select Case CLng(FrequencyNumber)
              Case 1
3                 sParseFrequencyNumber = "Annual"
4             Case 2
5                 sParseFrequencyNumber = "Semi annual"
6             Case 4
7                 sParseFrequencyNumber = "Quarterly"
8             Case 12
9                 sParseFrequencyNumber = "Monthly"
10            Case Else
11                If ThrowOnError Then
12                    Throw ErrString
13                Else
14                    sParseFrequencyNumber = "#" + ErrString + "!"
15                End If
16        End Select
17        Exit Function
ErrHandler:
18        Throw "#sParseFrequencyNumber (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : JuliaTradesToPortfolioTrades
' Author    : Hermione Glyn, Philip Swannell
' Date      : 03-Aug-2016
' Purpose   : Reverse of PortfolioTradesToJuliaTrades. Takes trades in Julia format and makes them
'             suitable for the Portfolio sheet of XVAFrontEnd.
'---------------------------------------------------------------------------------------
Function JuliaTradesToPortfolioTrades(InTrades As Variant, Numeraire As String)

          Dim CopyOfErr As String
          Dim i As Long
          Dim inTradeHeaders As Variant
          Dim j As Long
          Dim nbTrades As Long
          Dim OutTrades() As Variant
          Dim ThisTradeID As String
          Dim ThisVF As String
          
          Dim nIn_CashflowAmounts As Long
          Dim nIn_CashflowDates As Long
          Dim nIn_Counterparty As Long
          Dim nIn_Coupon As Long
          Dim nIn_Currency As Long
          Dim nIn_Dates As Long
          Dim nIn_EndDate As Long
          Dim nIn_FixedBDC As Long
          Dim nIn_FixedDCT As Long
          Dim nIn_FixedFrequency As Long
          Dim nIn_FloatingBDC As Long
          Dim nIn_FloatingDCT As Long
          Dim nIn_FloatingFrequency As Long
          Dim nIn_IsCall As Long
          Dim nIn_Notional As Long
          Dim nIn_Notionals As Long
          Dim nIn_PayAmortNotionals As Long
          Dim nIn_PayBDC As Long
          Dim nIn_PayCoupon As Long
          Dim nIn_PayCurrency As Long
          Dim nIn_PayDCT As Long
          Dim nIn_PayFrequency As Long
          Dim nIn_PayIndex As Long
          Dim nIn_PayLegType As Long
          Dim nIn_PayNotional As Long
          Dim nIn_PayNotionals As Long
          Dim nIn_ReceiveAmortNotionals As Long
          Dim nIn_ReceiveBDC As Long
          Dim nIn_ReceiveCoupon As Long
          Dim nIn_ReceiveCurrency As Long
          Dim nIn_ReceiveDCT As Long
          Dim nIn_ReceiveFrequency As Long
          Dim nIn_ReceiveIndex As Long
          Dim nIn_ReceiveLegType As Long
          Dim nIn_ReceiveNotional As Long
          Dim nIn_ReceiveNotionals As Long
          Dim nIn_StartDate As Long
          Dim nIn_Strike As Long
          Dim nIn_Strikes As Long
          Dim nIn_TradeID As Long
          Dim nIn_ValuationFunction As Long
          
1         On Error GoTo ErrHandler

2         If sNRows(InTrades) <= 1 Then
3             nbTrades = 0
4             JuliaTradesToPortfolioTrades = Empty
5             Exit Function
6         Else
7             nbTrades = sNRows(InTrades) - 1
8         End If

9         inTradeHeaders = sArrayTranspose(sSubArray(InTrades, 1, 1, 1))

10        nIn_TradeID = SafeMatch(inTradeHeaders, "TradeID")
11        nIn_ValuationFunction = SafeMatch(inTradeHeaders, "ValuationFunction")
12        nIn_Counterparty = SafeMatch(inTradeHeaders, "Counterparty")
13        nIn_StartDate = SafeMatch(inTradeHeaders, "StartDate")
14        nIn_EndDate = SafeMatch(inTradeHeaders, "EndDate")
15        nIn_Currency = SafeMatch(inTradeHeaders, "Currency")
16        nIn_ReceiveCurrency = SafeMatch(inTradeHeaders, "ReceiveCurrency")
17        nIn_PayCurrency = SafeMatch(inTradeHeaders, "PayCurrency")
18        nIn_Notional = SafeMatch(inTradeHeaders, "Notional")
19        nIn_ReceiveNotional = SafeMatch(inTradeHeaders, "ReceiveNotional")
20        nIn_PayNotional = SafeMatch(inTradeHeaders, "PayNotional")
21        nIn_ReceiveAmortNotionals = SafeMatch(inTradeHeaders, "ReceiveAmortNotionals")
22        nIn_PayAmortNotionals = SafeMatch(inTradeHeaders, "PayAmortNotionals")
23        nIn_Coupon = SafeMatch(inTradeHeaders, "Coupon")
24        nIn_ReceiveCoupon = SafeMatch(inTradeHeaders, "ReceiveCoupon")
25        nIn_PayCoupon = SafeMatch(inTradeHeaders, "PayCoupon")
26        nIn_Strike = SafeMatch(inTradeHeaders, "Strike")
27        nIn_ReceiveFrequency = SafeMatch(inTradeHeaders, "ReceiveFrequency")
28        nIn_PayFrequency = SafeMatch(inTradeHeaders, "PayFrequency")
29        nIn_FixedFrequency = SafeMatch(inTradeHeaders, "FixedFrequency")
30        nIn_FloatingFrequency = SafeMatch(inTradeHeaders, "FloatingFrequency")
31        nIn_ReceiveDCT = SafeMatch(inTradeHeaders, "ReceiveDCT")
32        nIn_PayDCT = SafeMatch(inTradeHeaders, "PayDCT")
33        nIn_FixedDCT = SafeMatch(inTradeHeaders, "FixedDCT")
34        nIn_FloatingDCT = SafeMatch(inTradeHeaders, "FloatingDCT")
35        nIn_ReceiveBDC = SafeMatch(inTradeHeaders, "ReceiveBDC")
36        nIn_PayBDC = SafeMatch(inTradeHeaders, "PayBDC")
37        nIn_FixedBDC = SafeMatch(inTradeHeaders, "FixedBDC")
38        nIn_FloatingBDC = SafeMatch(inTradeHeaders, "FloatingBDC")
39        nIn_IsCall = SafeMatch(inTradeHeaders, "IsCall")
40        nIn_CashflowDates = SafeMatch(inTradeHeaders, "CashflowDates")
41        nIn_CashflowAmounts = SafeMatch(inTradeHeaders, "CashflowAmounts")
42        nIn_Dates = SafeMatch(inTradeHeaders, "Dates")
43        nIn_Notionals = SafeMatch(inTradeHeaders, "Notionals")
44        nIn_Strikes = SafeMatch(inTradeHeaders, "Strikes")
45        nIn_ReceiveNotionals = SafeMatch(inTradeHeaders, "ReceiveNotionals")
46        nIn_PayNotionals = SafeMatch(inTradeHeaders, "PayNotionals")
47        nIn_ReceiveLegType = SafeMatch(inTradeHeaders, "ReceiveLegType")
48        nIn_PayLegType = SafeMatch(inTradeHeaders, "PayLegType")
49        nIn_ReceiveIndex = SafeMatch(inTradeHeaders, "ReceiveIndex")
50        nIn_PayIndex = SafeMatch(inTradeHeaders, "PayIndex")

51        ReDim OutTrades(1 To nbTrades, 1 To 19)

52        For i = 2 To nbTrades + 1
53            ThisTradeID = InTrades(i, nIn_TradeID)
54            j = i - 1
55            OutTrades(j, gCN_TradeID) = InTrades(i, nIn_TradeID)
56            CopyItem OutTrades, j, gCN_Counterparty, InTrades, i, nIn_Counterparty, vbString, ThisTradeID, "Counterparty"
57            ThisVF = InTrades(i, nIn_ValuationFunction)
58            OutTrades(j, gCN_TradeType) = ThisVF

59            Select Case ThisVF

                  Case "CrossCurrencySwap", "InterestRateSwap"
60                    CopyItem OutTrades, j, gCN_StartDate, InTrades, i, nIn_StartDate, vbDouble, ThisTradeID, "Start Date"
61                    CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "End Date"
62                    CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_PayCurrency, vbString, ThisTradeID, "Pay Currency"
63                    CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_ReceiveCurrency, vbString, ThisTradeID, "Pay Currency"
64                    CopyItem OutTrades, j, gCN_Rate1, InTrades, i, nIn_ReceiveCoupon, vbDouble, ThisTradeID, "Receive Coupon"
65                    CopyItem OutTrades, j, gCN_Rate2, InTrades, i, nIn_PayCoupon, vbDouble, ThisTradeID, "Pay Coupon"
66                    OutTrades(j, gCN_Freq1) = sParseFrequencyNumber(CStr(InTrades(i, nIn_ReceiveFrequency)), True)
67                    OutTrades(j, gCN_Freq2) = sParseFrequencyNumber(CStr(InTrades(i, nIn_PayFrequency)), True)
68                    OutTrades(j, gCN_DCT1) = sParseDCT(CStr(InTrades(i, nIn_ReceiveDCT)), False, True)    'IsFloating argument passed as False since ParseDCT only uses that to throw errors if the DCT "looks unlikely"
69                    OutTrades(j, gCN_DCT2) = sParseDCT(CStr(InTrades(i, nIn_PayDCT)), False, True)    'ditto
70                    OutTrades(j, gCN_BDC1) = ValidateBDC(CStr(InTrades(i, nIn_ReceiveBDC)), True)
71                    OutTrades(j, gCN_BDC2) = ValidateBDC(CStr(InTrades(i, nIn_PayBDC)), True)

72                    If isNonTrivialString(InTrades(i, nIn_ReceiveAmortNotionals)) Then     'Trade is amortising
73                        OutTrades(j, gCN_Notional1) = InTrades(i, nIn_ReceiveAmortNotionals)
74                        OutTrades(j, gCN_Notional2) = Replace(InTrades(i, nIn_PayAmortNotionals), "-", "")
75                    Else
76                        CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_ReceiveNotional, vbDouble, ThisTradeID, "Receive Notional"
77                        CopyItem OutTrades, j, gCN_Notional2, InTrades, i, nIn_PayNotional, vbDouble, ThisTradeID, "Pay Notional"
78                        OutTrades(j, gCN_Notional2) = Abs(OutTrades(j, gCN_Notional2))
79                    End If
80                    CopyItem OutTrades, j, gCN_LegType1, InTrades, i, nIn_ReceiveIndex, vbString, ThisTradeID, "Receive Index"
81                    CopyItem OutTrades, j, gCN_LegType2, InTrades, i, nIn_PayIndex, vbString, ThisTradeID, "Pay Index"

82                Case "FxForward"
83                    If SameSign(CDbl(InTrades(i, nIn_ReceiveNotional)), CDbl(InTrades(i, nIn_PayNotional))) Then Throw "ReceiveNotional and PayNotional must not be the same sign"
84                    CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "Maturity Date"
85                    If InTrades(i, nIn_ReceiveNotional) > 0 Then
86                        CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_ReceiveCurrency, vbString, ThisTradeID, "Receive Currency"
87                        CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_PayCurrency, vbString, ThisTradeID, "Pay Currency"
88                        CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_ReceiveNotional, vbDouble, ThisTradeID, "Receive Notional"
89                        CopyItem OutTrades, j, gCN_Notional2, InTrades, i, nIn_PayNotional, vbDouble, ThisTradeID, "Pay Notional"
                          'Flip sign on Pay side
90                        OutTrades(j, gCN_Notional2) = -OutTrades(j, gCN_Notional2)
91                    Else
                          'In the conventions of the Portfolio sheet notionals must be positive, so we have to switch the pay side and receive side
92                        CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_PayCurrency, vbString, ThisTradeID, "Receive Currency"
93                        CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_ReceiveCurrency, vbString, ThisTradeID, "Pay Currency"
94                        CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_PayNotional, vbDouble, ThisTradeID, "Receive Notional"
95                        CopyItem OutTrades, j, gCN_Notional2, InTrades, i, nIn_ReceiveNotional, vbDouble, ThisTradeID, "Pay Notional"
96                        OutTrades(j, gCN_Notional2) = -OutTrades(j, gCN_Notional2)
97                    End If

98                Case "FxForwardStrip"
                      'If SameSign(CDbl(InTrades(i, nIn_ReceiveNotional)), CDbl(InTrades(i, nIn_PayNotional))) Then Throw "ReceiveNotional and PayNotional must not be the same sign"
99                    CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_Dates, vbString, ThisTradeID, "Maturity Date"
100                   If Left(InTrades(i, nIn_ReceiveNotionals), 1) <> "-" Then    'ReceiveNotionals are positive
101                       CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_ReceiveCurrency, vbString, ThisTradeID, "Receive Currency"
102                       CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_PayCurrency, vbString, ThisTradeID, "Pay Currency"
103                       CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_ReceiveNotionals, vbString, ThisTradeID, "Receive Notional"
104                       CopyItem OutTrades, j, gCN_Notional2, InTrades, i, nIn_PayNotionals, vbString, ThisTradeID, "Pay Notional"
                          'Flip sign on Pay side
105                       OutTrades(j, gCN_Notional2) = FlipSCDS(CStr(OutTrades(j, gCN_Notional2)))
106                   Else
                          'In the conventions of the Portfolio sheet notionals must be positive, so we have to switch the pay side and receive side
107                       CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_PayCurrency, vbString, ThisTradeID, "Receive Currency"
108                       CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_ReceiveCurrency, vbString, ThisTradeID, "Pay Currency"
109                       CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_PayNotionals, vbString, ThisTradeID, "Receive Notional"
110                       CopyItem OutTrades, j, gCN_Notional2, InTrades, i, nIn_ReceiveNotionals, vbString, ThisTradeID, "Pay Notional"
111                       OutTrades(j, gCN_Notional2) = FlipSCDS(CStr(OutTrades(j, gCN_Notional2)))
112                   End If
113               Case "FxOption"
                      'In R format we consider all Fx options as being an option on the Currency versus the Numeraire currency
114                   CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "Maturity Date"
                      'Currencies...
115                   OutTrades(j, gCN_Ccy1) = Numeraire
116                   CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
                      'Notionals...
117                   If VarType(InTrades(i, nIn_Notional)) <> vbDouble Then Throw "Notional must be a number"
118                   If VarType(InTrades(i, nIn_Strike)) <> vbDouble Then Throw "Strike must be a positive number"
119                   If InTrades(i, nIn_Strike) <= 0 Then Throw "Strike must be a positive number"
120                   OutTrades(j, gCN_Notional2) = Abs(InTrades(i, nIn_Notional))
121                   OutTrades(j, gCN_Notional1) = Abs(InTrades(i, nIn_Notional)) * InTrades(i, nIn_Strike)
                      'IsFixed1 ? i.e. string declaring whether we are long or short option and whether it's a put or call. _
                       In R we think of puts and calls on the non-numeraire currency whereas on the Portfolio sheet we have _
                       put the Numeraire on the left (ReceiveCurrency or Ccy 1) so we describe whether it's a put or call _
                       on the Numeraire i.e. flip Call to Put and Vice Versa
122                   If VarType(InTrades(i, nIn_IsCall)) <> vbBoolean Then Throw "IsCall must be TRUE or FALSE"
                      Dim optDesc As String
123                   If InTrades(i, nIn_Notional) > 0 Then
124                       If InTrades(i, nIn_IsCall) Then
125                           optDesc = "BuyPut"    'flipped - see explanation above
126                       Else
127                           optDesc = "BuyCall"
128                       End If
129                   Else
130                       If InTrades(i, nIn_IsCall) Then
131                           optDesc = "SellPut"    'flipped - see explanation above
132                       Else
133                           optDesc = "SellCall"
134                       End If
135                   End If
136                   OutTrades(j, gCN_LegType1) = optDesc
137               Case "FxOptionStrip"
                      'In R format we consider all Fx options as being an option on the Currency versus the Numeraire currency
138                   If VarType(InTrades(i, nIn_Dates)) <> vbString Then Throw "Dates must be a semi-colon delimited number"
139                   CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_Dates, vbString, ThisTradeID, "Dates"
                      'Currencies...
140                   OutTrades(j, gCN_Ccy1) = Numeraire
141                   CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
                      'Notionals...
142                   If VarType(InTrades(i, nIn_Notionals)) <> vbString Then Throw "Notionals must be a semi-colon delimited string"
143                   If VarType(InTrades(i, nIn_Strikes)) <> vbString Then Throw "Strikes must be a semi-colon delimited string"
                      '  If InTrades(i, nIn_Strike) <= 0 Then Throw "Strike must be a positive number"
144                   OutTrades(j, gCN_Notional2) = Replace(InTrades(i, nIn_Notionals), "-", "")    'Replace(,"-","") takes absolute values
145                   OutTrades(j, gCN_Notional1) = MultiplySCDS(Replace(InTrades(i, nIn_Notionals), "-", ""), CStr(InTrades(i, nIn_Strikes)))
                      'IsFixed1 ? i.e. string declaring whether we are long or short option and whether it's a put or call. _
                       In R we think of puts and calls on the non-numeraire currency whereas on the Portfolio sheet we have _
                       put the Numeraire on the left (ReceiveCurrency or Ccy 1) so we describe whether it's a put or call _
                       on the Numeraire i.e. flip Call to Put and Vice Versa
146                   If VarType(InTrades(i, nIn_IsCall)) <> vbBoolean Then Throw "IsCall must be TRUE or FALSE"
147                   If Left(InTrades(i, nIn_Notionals), 1) <> "-" Then
148                       If InTrades(i, nIn_IsCall) Then
149                           optDesc = "BuyPut"    'flipped - see explanation above
150                       Else
151                           optDesc = "BuyCall"
152                       End If
153                   Else
154                       If InTrades(i, nIn_IsCall) Then
155                           optDesc = "SellPut"    'flipped - see explanation above
156                       Else
157                           optDesc = "SellCall"
158                       End If
159                   End If
160                   OutTrades(j, gCN_LegType1) = optDesc

161               Case "Swaption"
162                   CopyItem OutTrades, j, gCN_StartDate, InTrades, i, nIn_StartDate, vbDouble, ThisTradeID, "Start Date"
163                   CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "Maturity Date"
164                   CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
165                   CopyItem OutTrades, j, gCN_Ccy2, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
166                   If VarType(InTrades(i, nIn_Notional)) <> vbDouble Then Throw "Notional must be a Number"
167                   Select Case CStr(InTrades(i, nIn_IsCall))
                          Case "True"
168                           If InTrades(i, nIn_Notional) > 0 Then
169                               OutTrades(j, gCN_LegType1) = "BuyPayers"
170                               OutTrades(j, gCN_Notional1) = InTrades(i, nIn_Notional)  'Notional, have already tested for is number
171                           Else
172                               OutTrades(j, gCN_LegType1) = "SellPayers"
173                               OutTrades(j, gCN_Notional1) = InTrades(i, nIn_Notional) * -1  'Notional, have already tested for is number
174                           End If
175                       Case "False"
176                           If InTrades(i, nIn_Notional) > 0 Then
177                               OutTrades(j, gCN_LegType1) = "BuyReceivers"
178                               OutTrades(j, gCN_Notional1) = InTrades(i, nIn_Notional)  'Notional, have already tested for is number
179                           Else
180                               OutTrades(j, gCN_LegType1) = "SellReceivers"
181                               OutTrades(j, gCN_Notional1) = InTrades(i, nIn_Notional) * -1  'Notional, have already tested for is number
182                           End If
183                       Case Else
184                           Throw "For options, IsCall must have a True/False value."
185                   End Select
186                   CopyItem OutTrades, j, gCN_Rate1, InTrades, i, nIn_Strike, vbDouble, ThisTradeID, "Coupon"
187                   OutTrades(j, gCN_Freq1) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FixedFrequency)), True)    'FixedFrequency
188                   OutTrades(j, gCN_Freq2) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FloatingFrequency)), True)    'FloatingFrequency
189                   OutTrades(j, gCN_DCT1) = sParseDCT(CStr(InTrades(i, nIn_FixedDCT)), False, True)   'FixedDCT
190                   OutTrades(j, gCN_DCT2) = sParseDCT(CStr(InTrades(i, nIn_FloatingDCT)), True, True)   'FloatingDCT
191                   OutTrades(j, gCN_BDC1) = ValidateBDC(CStr(InTrades(i, nIn_FixedBDC)), True)    'FixedBDC
192                   OutTrades(j, gCN_BDC2) = ValidateBDC(CStr(InTrades(i, nIn_FloatingBDC)), True)    'FloatingBDC
193                   OutTrades(j, gCN_Notional2) = OutTrades(j, gCN_Notional1)
194               Case "CapFloor"
195                   CopyItem OutTrades, j, gCN_StartDate, InTrades, i, nIn_StartDate, vbDouble, ThisTradeID, "Start Date"
196                   CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "Maturity Date"
197                   CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
198                   Select Case CStr(InTrades(i, nIn_IsCall))
                          Case "True"
199                           If InTrades(i, nIn_Notional) > 0 Then
200                               CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_Notional, vbDouble, ThisTradeID, "Notional"
201                               OutTrades(j, gCN_LegType1) = "BuyCap"
202                           Else
203                               CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_Notional, vbDouble, ThisTradeID, "Notional"
204                               OutTrades(j, gCN_Notional1) = -OutTrades(j, gCN_Notional1)  'Flip sign
205                               OutTrades(j, gCN_LegType1) = "SellCap"
206                           End If
207                       Case "False"
208                           If InTrades(i, nIn_Notional) > 0 Then
209                               CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_Notional, vbDouble, ThisTradeID, "Notional"
210                               OutTrades(j, gCN_LegType1) = "BuyFloor"
211                           Else
212                               CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_Notional, vbDouble, ThisTradeID, "Notional"
213                               OutTrades(j, gCN_Notional1) = -OutTrades(j, gCN_Notional1)  'Flip sign
214                               OutTrades(j, gCN_LegType1) = "SellFloor"
215                           End If
216                       Case Else
217                           Throw "Unrecognised option style for trade " + ThisTradeID
218                   End Select
219                   If InTrades(i, nIn_ReceiveNotional) < 0 Then Throw "Notional must be positive or zero"
220                   CopyItem OutTrades, j, gCN_Rate1, InTrades, i, nIn_Strike, vbDouble, ThisTradeID, "Strike"
221                   OutTrades(j, gCN_Freq1) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FloatingFrequency)), True)
222                   OutTrades(j, gCN_DCT1) = sParseDCT(CStr(InTrades(i, nIn_FloatingDCT)), True, True)
223                   OutTrades(j, gCN_BDC1) = ValidateBDC(CStr(InTrades(i, nIn_FloatingBDC)), True)
224               Case "FixedCashflows"
                      ' ValidateFixedCashflows ThisTradeID, InTrades(i, nIn_EndDate), InTrades(i, nIn_ReceiveNotional)
225                   OutTrades(j, gCN_EndDate) = InTrades(i, nIn_CashflowDates)
226                   OutTrades(j, gCN_Notional1) = InTrades(i, nIn_CashflowAmounts)
227                   CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
228               Case "InflationZCSwap", "InflationYoYSwap"
229                   CopyItem OutTrades, j, gCN_StartDate, InTrades, i, nIn_StartDate, vbDouble, ThisTradeID, "Start Date"
230                   CopyItem OutTrades, j, gCN_EndDate, InTrades, i, nIn_EndDate, vbDouble, ThisTradeID, "Maturity Date"
231                   CopyItem OutTrades, j, gCN_Ccy1, InTrades, i, nIn_Currency, vbString, ThisTradeID, "Currency"
232                   CopyItem OutTrades, j, gCN_Notional1, InTrades, i, nIn_Notional, vbDouble, ThisTradeID, "Notional"
233                   OutTrades(j, gCN_Notional1) = Abs(OutTrades(j, gCN_Notional1))
234                   OutTrades(j, gCN_Notional2) = Abs(OutTrades(j, gCN_Notional1))
235                   If InTrades(i, nIn_Notional) >= 0 Then    ' Receive fixed, pay index
236                       OutTrades(j, gCN_LegType1) = "Fixed"
237                       OutTrades(j, gCN_LegType2) = "Index"
238                       CopyItem OutTrades, j, gCN_BDC1, InTrades, i, nIn_FixedBDC, vbString, ThisTradeID, "FixedBDC"
239                       CopyItem OutTrades, j, gCN_BDC2, InTrades, i, nIn_FloatingBDC, vbString, ThisTradeID, "FloatingBDC"
240                       CopyItem OutTrades, j, gCN_Rate1, InTrades, i, nIn_Coupon, vbDouble, ThisTradeID, "Coupon"
241                       If ThisVF = "InflationYoYSwap" Then
242                           OutTrades(j, gCN_Freq1) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FixedFrequency)), True)
243                           OutTrades(j, gCN_Freq2) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FloatingFrequency)), True)
244                       End If
245                   Else    ' Pay fixed, receive index
246                       OutTrades(j, gCN_LegType1) = "Index"
247                       OutTrades(j, gCN_LegType2) = "Fixed"
248                       CopyItem OutTrades, j, gCN_BDC1, InTrades, i, nIn_FloatingBDC, vbString, ThisTradeID, "FloatingBDC"
249                       CopyItem OutTrades, j, gCN_BDC2, InTrades, i, nIn_FixedBDC, vbString, ThisTradeID, "FixedBDC"
250                       CopyItem OutTrades, j, gCN_Rate2, InTrades, i, nIn_Coupon, vbDouble, ThisTradeID, "Coupon"
251                       If ThisVF = "InflationYoYSwap" Then
252                           OutTrades(j, gCN_Freq1) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FloatingFrequency)), True)
253                           OutTrades(j, gCN_Freq2) = sParseFrequencyNumber(CStr(InTrades(i, nIn_FixedFrequency)), True)
254                       End If
255                   End If
256               Case Else
257                   Throw "Unrecognised value in TradeType column: " + ThisVF
258           End Select
259       Next i

260       JuliaTradesToPortfolioTrades = OutTrades

261       Exit Function
ErrHandler:
262       CopyOfErr = Err.Description
263       If InStr(CopyOfErr, ThisTradeID) = 0 Then
264           CopyOfErr = "#JuliaTradesToPortfolioTrades (line " & CStr(Erl) + "): " + "(TradeID = " + CStr(ThisTradeID) + ") " & CopyOfErr & "!"
265       Else
266           CopyOfErr = "#JuliaTradesToPortfolioTrades (line " & CStr(Erl) + "): " & CopyOfErr & "!"
267       End If
268       Throw CopyOfErr
End Function

Private Function isNonTrivialString(x) As Boolean
1         If VarType(x) = vbString Then
2             If Len(x) > 0 Then
3                 isNonTrivialString = True
4             End If
5         End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : SameSign
' Author    : Philip Swannell
' Date      : 16-Dec-2016
' Purpose   : Returns true if x and y are either both positive or both negative
'---------------------------------------------------------------------------------------
Private Function SameSign(x As Double, y As Double) As Boolean
1         On Error GoTo ErrHandler
2         If x > 0 And y > 0 Then
3             SameSign = True
4         ElseIf x < 0 And y < 0 Then
5             SameSign = True
6         Else
7             SameSign = False
8         End If
9         Exit Function
ErrHandler:
10        Throw "#SameSign (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : PasteTradesToPortfolioSheet
' Author    : Philip Swannell
' Date      : 18 Dec 2016
' Purpose   : Pastes an array of data to the portfolio sheet. No checking that the data is valid.
'---------------------------------------------------------------------------------------
Function PasteTradesToPortfolioSheet(TradesPortfolioFormat As Variant, Optional SourceFileName As String, Optional Overwrite As Variant)
          Dim Append As Boolean
          Dim CopyOfErr As String
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim NumExistingTrades As Long
          Dim OldBCE As Boolean
          Dim Res As VbMsgBoxResult
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TargetRange

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         getTradesRange NumExistingTrades

5         If NumExistingTrades = 0 Then
6             Append = False
7         Else
8             If VarType(Overwrite) = vbBoolean Then
9                 Append = Not Overwrite
10            Else
11                Res = MsgBoxPlus("Do you want to" + vbLf + "a) Append the trades in the file the to the trades on the sheet; or" + vbLf + _
                      "b) Overwrite the trades on the sheet with the trades in the file?", _
                      vbYesNoCancel + vbQuestion + vbDefaultButton3, MsgBoxTitle(), "Append", "Overwrite", , , 310)
12                If Res = vbYes Then
13                    Append = True
14                ElseIf Res = vbNo Then
15                    Append = False
16                Else
17                    GoTo EarlyExit
18                End If
19            End If
20        End If

21        Set SUH = CreateScreenUpdateHandler()
22        Set ExSH = CreateExcelStateHandler(PreserveViewport:=True)
23        Application.DisplayAlerts = False
24        Set SPH = CreateSheetProtectionHandler(shPortfolio)
25        If Append = False Then
26            With getTradesRange(1)
27                .Resize(.Rows.Count + 1).Clear    ' make sure we clear out the "<Doubleclick to add trade>" label
28            End With
29            Set TargetRange = getTradesRange(1).Resize(sNRows(TradesPortfolioFormat), sNCols(TradesPortfolioFormat))
30        Else
31            With getTradesRange(NumExistingTrades)
32                Set TargetRange = .Cells(NumExistingTrades + 1, 1).Resize(sNRows(TradesPortfolioFormat), sNCols(TradesPortfolioFormat))
33            End With
34        End If
35        TargetRange.Value = sArrayExcelString(TradesPortfolioFormat)
36        FormatTradesRange
37        If Append Then
38            If TradeIDsNeedRepairing Then
39                RepairTradeIDs
40            End If
41        End If
42        RangeFromSheet(shPortfolio, "TradesFileName").Value = "'" + SourceFileName

43        FilterTradesRange
44        CalculatePortfolioSheet
45        shxVADashboard.Calculate
46        ResetSortButtons RangeFromSheet(shPortfolio, "PortfolioHeader").Offset(-2, 0).Resize(, RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Columns.Count), False, False

EarlyExit:
47        gBlockChangeEvent = OldBCE
48        Exit Function
ErrHandler:
49        CopyOfErr = "#PasteTradesToPortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
50        gBlockChangeEvent = OldBCE
51        Throw CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : TestTradeConversion
' Author    : Philip Swannell
' Date      : 23-Dec-2016
' Purpose   : Test harness for PortfolioTradesToJuliaTrades and JuliaTradesToPortfolioTrades being
'             correct inverses of one another
'---------------------------------------------------------------------------------------
Sub TestTradeConversion()
          Dim finalPortfolioTrades
          Dim NumTrades
          Dim origPortfolioTrades
          Dim Prompt As String
          Dim RTrades
          Const Title = "Test trade conversion"
          Dim Headers
          Dim RangeToFormat As Range
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Static SuppressInitialDialog As Boolean

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         origPortfolioTrades = getTradesRange(NumTrades).Resize(, 19).Value2
4         If NumTrades = 0 Then Exit Sub

5         If Not SuppressInitialDialog Then
6             Prompt = "This method takes the trades currently displayed, " & _
                  "converts them to the format required for valuation in Julia and then converts " & _
                  "them back to the format displayed on the Portfolio sheet." + vbLf + vbLf + _
                  "Such conversion back and forth should not change the trades. If changes do happen the differences will be displayed to aid debugging."
7             Prompt = Prompt + vbLf + vbLf + "Proceed?"
8             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, , , , , , "Don't show this message again.", SuppressInitialDialog) <> vbOK Then Exit Sub
9         End If

10        origPortfolioTrades = sArrayIf(sArrayEquals(origPortfolioTrades, "N/A"), Empty, origPortfolioTrades)

11        RTrades = PortfolioTradesToJuliaTrades(origPortfolioTrades, True, False)

12        finalPortfolioTrades = JuliaTradesToPortfolioTrades(RTrades, RangeFromMarketDataBook("Config", "Numeraire").Value)

13        If sArraysIdentical(origPortfolioTrades, finalPortfolioTrades) Then
14            Prompt = "All trades unchanged by conversion to Julia format and back to Portfolio format."
15            MsgBoxPlus Prompt, vbInformation, Title
16        Else
17            Headers = getTradesRange(NumTrades).Rows(-1).Resize(2, 19).Value
18            Prompt = "Detected that conversion of trades to Julia format and back to Portfolio format changed some of the trades."
19            g sDiffTwoArrays(sArrayStack(Headers, origPortfolioTrades), sArrayStack(Headers, finalPortfolioTrades)), ExMthdSpreadsheet
20            Set RangeToFormat = ActiveSheet.UsedRange
21            RangeToFormat.FormatConditions.Add Type:=xlExpression, Formula1:= _
                  "=IFS(ISNUMBER(A1),A1<>0,ISTEXT(A1),LEFT(A1,1)=""["")"
22            RangeToFormat.FormatConditions(RangeToFormat.FormatConditions.Count).SetFirstPriority
23            With RangeToFormat.FormatConditions(1).Interior
24                .PatternColorIndex = xlAutomatic
25                .Color = 65535
26                .TintAndShade = 0
27            End With
28            RangeToFormat.FormatConditions(1).StopIfTrue = False
29            MsgBoxPlus Prompt, vbExclamation, Title
30        End If

31        Exit Sub
ErrHandler:
32        SomethingWentWrong "#TestTradeConversion (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

