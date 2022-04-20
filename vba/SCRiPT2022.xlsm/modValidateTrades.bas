Attribute VB_Name = "modValidateTrades"
Option Explicit

'Constants for column order of trades on the Portfolio sheet
Public Const gCN_TradeID As Long = 1
Public Const gCN_TradeType As Long = 2
Public Const gCN_StartDate As Long = 3
Public Const gCN_EndDate As Long = 4
Public Const gCN_Ccy1 As Long = 5
Public Const gCN_Notional1 As Long = 6
Public Const gCN_Rate1 As Long = 7
Public Const gCN_LegType1 As Long = 8
Public Const gCN_Freq1 As Long = 9
Public Const gCN_DCT1 As Long = 10
Public Const gCN_BDC1 As Long = 11
Public Const gCN_Ccy2 As Long = 12
Public Const gCN_Notional2 As Long = 13
Public Const gCN_Rate2 As Long = 14
Public Const gCN_LegType2 As Long = 15
Public Const gCN_Freq2 As Long = 16
Public Const gCN_DCT2 As Long = 17
Public Const gCN_BDC2 As Long = 18
Public Const gCN_Counterparty As Long = 19

Sub TestValidateTrades()
1         On Error GoTo ErrHandler
          Dim t1 As Double
          Dim t2 As Double
2         t1 = sElapsedTime
3         ValidateTrades getTradesRange(0).Value2
4         t2 = sElapsedTime
5         Debug.Print t2 - t1
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#TestValidateTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ValidateTrades
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Throws error if there is manifest error in the trade data in the format that is
'             held on the Portfolio sheet. Better than getting cryptic error back from the Julia code...
'             TODO - this method has become unmanageable - better to have separate routines for each
'             trade type and this routine call the underlying methods?
'             Note also that this method duplicates validation in method PortfolioTradesToJuliaTrades. Belt and Braces perhaps.
'---------------------------------------------------------------------------------------
Function ValidateTrades(Trades, Optional ThrowOnError As Boolean = True, Optional ByRef ErrorMessage As String)
                   
          Dim Col_Ccy1 As Variant
          Dim Col_Ccy2 As Variant

          Dim BDCMatchIDs1
          Dim BDCMatchIDs2
          Dim Ccy1IsCurrency As Boolean
          Dim Ccy1IsInflation As Boolean
          Dim Ccy1MatchIDs
          Dim Ccy1MatchIDs_Inf
          Dim Ccy2IsCurrency As Boolean
          Dim Ccy2IsInflation As Boolean
          Dim Ccy2MatchIDs
          Dim Ccy2MatchIDs_Inf
          Dim CPMatchIDs
          Dim GoodCounterparties As Variant
          Dim GoodCurrencies As Variant
          Dim i As Long
          Dim NumTrades As Long
          Dim STK As SolumAddin.clsStacker
          Dim TheTradeIDs
          Dim ThisTradeID As Variant
          Dim ThisTradeType As String
          Dim vRes As String

1         On Error GoTo ErrHandler

2         CalculatePortfolioSheet

3         GoodCurrencies = sSortedArray(CurrenciesSupported(False, True))
4         OpenMarketWorkbook True, False

5         TheTradeIDs = sSubArray(Trades, 1, gCN_TradeID, , 1)
6         NumTrades = sNRows(TheTradeIDs)
7         If NumTrades <> sNRows(sRemoveDuplicates(TheTradeIDs)) Then Throw "Duplicate TradeIDs exist. Fix with" + vbLf + "Menu > Trades > Repair invalid Trade IDs", True

8         GoodCounterparties = CounterpartiesFromMarketBook(True)
9         CPMatchIDs = sMatch(sSubArray(Trades, 1, gCN_Counterparty, , 1), GoodCounterparties)

10        BDCMatchIDs1 = sMatch(sSubArray(Trades, 1, gCN_BDC1, , 1), SupportedBDCs())
11        BDCMatchIDs2 = sMatch(sSubArray(Trades, 1, gCN_BDC2, , 1), SupportedBDCs())
12        Col_Ccy1 = sSubArray(Trades, 1, gCN_Ccy1, , 1)
13        Col_Ccy2 = sSubArray(Trades, 1, gCN_Ccy2, , 1)
14        Ccy1MatchIDs = sMatch(Col_Ccy1, GoodCurrencies, True)
15        Ccy2MatchIDs = sMatch(Col_Ccy2, GoodCurrencies, True)
16        Ccy1MatchIDs_Inf = sMatch(Col_Ccy1, SupportedInflationIndices(), True)
17        Ccy2MatchIDs_Inf = sMatch(Col_Ccy2, SupportedInflationIndices(), True)

18        If NumTrades = 1 Then Force2DArrayRMulti CPMatchIDs, BDCMatchIDs1, BDCMatchIDs2, Ccy1MatchIDs, Ccy2MatchIDs, Ccy1MatchIDs_Inf

19        Set STK = CreateStacker()

20        For i = 1 To sNRows(Trades)
21            ThisTradeID = Trades(i, gCN_TradeID)

22            ThisTradeType = Trades(i, gCN_TradeType)
23            If VarType(ThisTradeID) <> vbString Then
24                ThisTradeID = CStr(ThisTradeID)
25                STK.StackData "TradeID must be text but for the " + OneToFirst(i) + " trade it's not text"
26            End If

27            Ccy1IsInflation = False
28            Ccy2IsInflation = False
29            Select Case ThisTradeType
                  Case "InterestRateSwap"
30                    Ccy1IsCurrency = True
31                    Ccy2IsCurrency = True
32                Case "CrossCurrencySwap"
33                    Ccy1IsCurrency = True
34                    Ccy2IsCurrency = True
35                Case "FxForward"
36                    Ccy1IsCurrency = True
37                    Ccy2IsCurrency = True
38                Case "FxOption"
39                    Ccy1IsCurrency = True
40                    Ccy2IsCurrency = True
41                Case "CapFloor"
42                    Ccy1IsCurrency = True
43                    Ccy2IsCurrency = False
44                Case "Swaption"
45                    Ccy1IsCurrency = True
46                    Ccy2IsCurrency = True
47                Case "FixedCashflows"
48                    Ccy1IsCurrency = True
49                    Ccy2IsCurrency = False
50                Case "FxOptionStrip"
51                    Ccy1IsCurrency = True
52                    Ccy2IsCurrency = True
53                Case "FxForwardStrip"
54                    Ccy1IsCurrency = True
55                    Ccy2IsCurrency = True
56                Case "InflationZCSwap"
57                    Ccy1IsCurrency = False
58                    Ccy2IsCurrency = False
59                    Ccy1IsInflation = True
60                Case "InflationYoYSwap"
61                    If Trades(i, gCN_LegType1) = "Index" Then
62                        Ccy1IsInflation = True
63                        Ccy1IsCurrency = False
64                    Else
65                        Ccy1IsInflation = False
66                        Ccy1IsCurrency = True
67                    End If
68                    If Trades(i, gCN_LegType2) = "Index" Then
69                        Ccy2IsInflation = True
70                        Ccy2IsCurrency = False
71                    Else
72                        Ccy2IsInflation = False
73                        Ccy2IsCurrency = True
74                    End If
75                Case Else
76                    STK.StackData "Unrecognised TradeType '" + ThisTradeType + "'for trade " + ThisTradeID
77            End Select
78            If Not IsWholeNumber(Trades(i, gCN_StartDate)) Then
79                If ThisTradeType <> "FxForward" And ThisTradeType <> "FxOption" And ThisTradeType <> "FixedCashflows" And Right(ThisTradeType, 5) <> "Strip" Then
80                    If IsNumber(Trades(i, gCN_StartDate)) Then
81                        STK.StackData "StartDate must be a whole number but is not for trade " + ThisTradeID
82                    Else
83                        STK.StackData "StartDate must be a number but is not for trade " + ThisTradeID
84                    End If
85                End If
86            End If
87            If Not IsNumber(Trades(i, gCN_EndDate)) Then
88                If ThisTradeType <> "FixedCashflows" And Right(ThisTradeType, 5) <> "Strip" Then
89                    If IsNumber(Trades(i, gCN_EndDate)) Then
90                        STK.StackData "EndDate must be a whole number but is not for trade " + ThisTradeID
91                    Else
92                        STK.StackData "EndDate must be a number but is not for trade " + ThisTradeID
93                    End If
94                End If
95            End If
96            If Right(ThisTradeType, 4) = "Swap" Or ThisTradeType = "CapFloor" Or ThisTradeType = "Swaption" Then
97                If VarType(Trades(i, gCN_EndDate)) <> vbDouble Then
98                    STK.StackData "Invalid EndDate for trade " + ThisTradeID
99                ElseIf VarType(Trades(i, gCN_StartDate)) <> vbDouble Then
100                   STK.StackData "Invalid StartDate for trade " + ThisTradeID
101               ElseIf Trades(i, gCN_EndDate) <= Trades(i, gCN_StartDate) Then
102                   STK.StackData "EndDate must be after StartDate but is not for trade " + ThisTradeID
103               End If
104           End If

105           If Ccy1IsInflation Then
106               If Not IsNumber(Ccy1MatchIDs_Inf(i, 1)) Then
107                   STK.StackData "Unrecognised inflation index '" + CStr(Col_Ccy1(i, 1)) + "' (in Ccy 1 column) for trade " + ThisTradeID
108               End If
109           ElseIf Ccy1IsCurrency Then
110               If Not IsNumber(Ccy1MatchIDs(i, 1)) Then
111                   STK.StackData "Unrecognised Ccy 1 '" + CStr(Col_Ccy1(i, 1)) + "' for trade " + ThisTradeID
112               End If
113           End If

114           If Not IsNumber(Trades(i, gCN_Notional1)) Then
115               If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Then
116                   If Not IsValidNotional(Trades(i, gCN_Notional1)) Then
117                       STK.StackData "Invalid Notional 1 for trade " + ThisTradeID + ". It should be be a positive number or a semi-colon delimited list of positive numbers"
118                   End If
119               ElseIf ThisTradeType <> "FixedCashflows" And Right(ThisTradeType, 5) <> "Strip" Then
120                   STK.StackData "Notional1 must be a number but is not a number for trade " + ThisTradeID
121               End If
122           ElseIf Trades(i, gCN_Notional1) < 0 Then
123               STK.StackData "Notional1 must be positive or zero, but it's negative for " + ThisTradeID
124           End If

125           If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Or ThisTradeType = "CapFloor" Or ThisTradeType = "Swaption" Then
126               If Not IsNumber(Trades(i, gCN_Rate1)) Then
127                   STK.StackData "Rate1 must be a number but is not a number for trade " + ThisTradeID
128               End If
129               Select Case LCase(CStr(Trades(i, gCN_Freq1)))
                      Case "annual", "semi annual", "quarterly", "monthly", "a", "s", "q", "m"
130                   Case Else
131                       STK.StackData "Freq1 is invalid for trade " + ThisTradeID
132               End Select
133               vRes = sParseDCT(CStr(Trades(i, gCN_DCT1)), CStr(Trades(i, gCN_LegType1)) = "Floating" Or ThisTradeType = "CapFloor", False)
134               If Left(vRes, 1) = "#" Then
135                   STK.StackData "DCT 1 is invalid for trade " + ThisTradeID + " " + vRes
136               End If
137               If Not IsNumber(BDCMatchIDs1(i, 1)) Then
138                   STK.StackData "BDC 1 is invalid for trade " + ThisTradeID
139               End If
140           End If

141           If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Then
142               Select Case CStr(Trades(i, gCN_LegType1))
                      Case "Fixed", "Libor", "OIS"
143                   Case Else
144                       STK.StackData "Invalid Leg Type 1 for trade " + ThisTradeID + ". It must must be a either Fixed, Libor or OIS"
145               End Select
146           End If

147           If Ccy2IsInflation Then
148               If Not IsNumber(Ccy2MatchIDs_Inf(i, 1)) Then
149                   STK.StackData "Unrecognised inflation index '" + CStr(Col_Ccy2(i, 1)) + "' (in Ccy 2 column) for trade " + ThisTradeID
150               End If
151           ElseIf Ccy2IsCurrency Then
152               If Not IsNumber(Ccy2MatchIDs(i, 1)) Then
153                   STK.StackData "Unrecognised Ccy 2 '" + CStr(Col_Ccy2(i, 1)) + "' for trade " + ThisTradeID
154               End If
155           End If

156           If ThisTradeType = "InterestRateSwap" Then
157               If Trades(i, gCN_Ccy1) <> Trades(i, gCN_Ccy2) Then
158                   STK.StackData "Ccy1 and Ccy2 must be the same for InterestRateSwaps, but they're different for trade " + ThisTradeID
159               End If
160           End If
161           If Not IsNumber(Trades(i, gCN_Notional2)) Then
162               If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Then
163                   If Not IsValidNotional(Trades(i, gCN_Notional2)) Then
164                       STK.StackData "Invalid Notional 2 for trade " + ThisTradeID + ". It should be be a positive number or a semi-colon delimited list of positive numbers"
165                   End If
166               ElseIf ThisTradeType <> "CapFloor" And ThisTradeType <> "FixedCashflows" And Right(ThisTradeType, 5) <> "Strip" Then
167                   STK.StackData "Notional 2 must be a number but is not a number for trade " + ThisTradeID
168               End If
169           ElseIf Trades(i, gCN_Notional2) < 0 Then
170               STK.StackData "Notional2 must be positive or zero, but it's negative for " + ThisTradeID
171           End If
172           If ThisTradeType = "InterestRateSwap" Then
173               If Trades(i, gCN_Notional1) <> Trades(i, gCN_Notional2) Then
174                   If IsNumber(Trades(i, gCN_Notional1)) Or IsNumber(Trades(i, gCN_Notional2)) Then
175                       STK.StackData "Notional 1 and Notional 2 must be the same for InterestRateSwaps but are different at row " + CStr(i) + " TradeID = '" & ThisTradeID & "'"
176                   End If
177               End If
178           End If
179           If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Then
180               If Not IsNumber(Trades(i, gCN_Rate2)) Then
181                   STK.StackData "Rate2 must be a number but is not a number for trade " + ThisTradeID
182               End If
183               Select Case CStr(Trades(i, gCN_LegType2))
                      Case "Fixed", "Libor", "OIS"
184                   Case Else
185                       STK.StackData "Invalid Leg Type 2 for trade " + ThisTradeID + ". It must must be a either Fixed, Libor or OIS"
186               End Select
187           End If
188           If ThisTradeType = "InterestRateSwap" Or ThisTradeType = "CrossCurrencySwap" Or ThisTradeType = "Swaption" Then
189               Select Case LCase(CStr(Trades(i, gCN_Freq2)))
                      Case "annual", "semi annual", "quarterly", "monthly", "a", "s", "q", "m"
190                   Case Else
191                       STK.StackData "Freq2 is invalid for trade " + ThisTradeID
192               End Select
193               vRes = sParseDCT(CStr(Trades(i, gCN_DCT2)), CStr(Trades(i, gCN_LegType2)) = "Floating", False)
194               If Left(vRes, 1) = "#" Then
195                   STK.StackData "DCT 2 is invalid for trade " + ThisTradeID + " " + vRes
196               End If
197               If Not IsNumber(BDCMatchIDs2(i, 1)) Then
198                   STK.StackData "BDC 2 is invalid for trade " + ThisTradeID
199               End If
200           End If

201           If ThisTradeType = "FxOption" Or ThisTradeType = "FxOptionStrip" Then
202               Select Case CStr(Trades(i, gCN_LegType1))
                      Case "BuyPut", "BuyCall", "SellPut", "SellCall"
203                   Case Else
204                       STK.StackData "Option style (in 'Is Fixed? 1' column) is invalid for trade " + ThisTradeID
205               End Select
206           ElseIf ThisTradeType = "CapFloor" Then
207               Select Case CStr(Trades(i, gCN_LegType1))
                      Case "BuyCap", "SellCap", "BuyFloor", "SellFloor"
208                   Case Else
209                       STK.StackData "Option style (in 'Is Fixed? 1' column) is invalid for trade " + ThisTradeID
210               End Select
211           ElseIf ThisTradeType = "Swaption" Then
212               Select Case CStr(Trades(i, gCN_LegType1))
                      Case "BuyReceivers", "SellReceivers", "BuyPayers", "SellPayers"
213                   Case Else
214                       STK.StackData "Option style (in 'Is Fixed? 1' column) is invalid for trade " + ThisTradeID
215               End Select
216           ElseIf ThisTradeType = "FixedCashflows" Then
217               ValidateFixedCashflows CStr(ThisTradeID), Trades(i, gCN_EndDate), Trades(i, gCN_Notional1), STK
218           ElseIf ThisTradeType = "InflationYoYSwap" Then
219               Select Case CStr(Trades(i, gCN_LegType1)) & "|" & CStr(Trades(i, gCN_LegType2))
                      Case "Index|Fixed", "Index|Floating", "Fixed|Index", "Floating|Index"
220                   Case Else
221                       STK.StackData "Invalid values for 'Is Fixed? 1' and 'Is Fixed? 2' one must read 'Index' and the other either 'Fixed' or 'Floating' but that's not the case for Trade ID " + ThisTradeID
222               End Select
223           End If
224           If Right(ThisTradeType, 5) = "Strip" Then
225               ValidateFxStrip CStr(ThisTradeID), Trades(i, gCN_EndDate), Trades(i, gCN_Notional1), Trades(i, gCN_Notional2), ThisTradeType = "FxOptionStrip", STK
226           End If

227           If VarType(Trades(i, gCN_Counterparty)) <> vbString Then
228               STK.StackData "Invalid Counterparty for trade " + ThisTradeID
229           ElseIf Trades(i, gCN_Counterparty) = "??" Then
230               STK.StackData "Invalid Counterparty for trade " + ThisTradeID
231           ElseIf Not IsNumber(CPMatchIDs(i, 1)) Then
232               STK.StackData "Invalid Counterparty for trade " + ThisTradeID + " - valid Counterparties are held on the Credit sheet of the market data workbook"
233           End If

234       Next i

          Dim NumErrors
          Dim TheErrors
          Const MaxNumErrorsToShow As Long = 50

235       TheErrors = STK.report

236       If Not sArraysIdentical(TheErrors, "#Nothing to report!") Then
237           NumErrors = sNRows(TheErrors)
238           ErrorMessage = "Found " + IIf(NumErrors = 1, " an error in the trade data:", CStr(NumErrors) + " errors in the trade data:")
239           If NumErrors = 1 Then
240               ErrorMessage = ErrorMessage + vbLf + TheErrors(1, 1)
241           ElseIf NumErrors > MaxNumErrorsToShow Then
242               ErrorMessage = ErrorMessage + vbLf + "Here are the first " + CStr(MaxNumErrorsToShow) + " errors found:" + vbLf _
                      + sConcatenateStrings(sSubArray(TheErrors, 1, 1, MaxNumErrorsToShow), vbLf)
243           Else
244               ErrorMessage = ErrorMessage + vbLf + _
                      sConcatenateStrings(TheErrors, vbLf)
245           End If
246           If ThrowOnError Then
247               Throw CStr(ErrorMessage), True
248           End If
249       End If

250       Exit Function
ErrHandler:
251       Throw "#ValidateTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : OneToFirst
' Author    : Philip Swannell
' Date      : 29-Feb-2016
' Purpose   : Translates cardinal to ordinal
'---------------------------------------------------------------------------------------
Private Function OneToFirst(i As Long) As String
          Dim Res As String
1         Res = CStr(i)
2         Select Case i Mod 100
              Case 11, 12, 13
3                 Res = Res + "th"
4             Case Else
5                 Select Case i Mod 10
                      Case 1
6                         Res = Res + "st"
7                     Case 2
8                         Res = Res + "nd"
9                     Case 3
10                        Res = Res + "rd"
11                    Case Else
12                        Res = Res + "th"
13                End Select
14        End Select
15        OneToFirst = Res
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsGoodDateString
' Author    : Philip Swannell
' Date      : 15-Mar-2016
' Purpose   :
'---------------------------------------------------------------------------------------
Private Function IsGoodDateString(DateString As String) As Boolean
1         On Error GoTo ErrHandler
2         If (LCase(Format(CDate(DateString), "dd-mmm-yyyy")) = LCase(DateString)) Then
3             IsGoodDateString = True
4         ElseIf (LCase(Format(CDate(DateString), "d-mmm-yyyy")) = LCase(DateString)) Then
5             IsGoodDateString = True
6         Else
7             IsGoodDateString = False
8         End If
9         Exit Function
ErrHandler:
10        IsGoodDateString = False
End Function

Function StringCanBeCastToDouble(NumberString As String) As Boolean
          Dim D As Double
1         On Error GoTo ErrHandler
2         D = CDbl(NumberString)
3         StringCanBeCastToDouble = True
4         Exit Function
ErrHandler:
5         StringCanBeCastToDouble = False
End Function

Function StringCanBeCastToDate(DateString As String) As Boolean
          Dim D As Date
1         On Error GoTo ErrHandler
2         D = CDate(DateString)
3         StringCanBeCastToDate = True
4         Exit Function
ErrHandler:
5         StringCanBeCastToDate = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateFixedCashflows
' Author    : Philip Swannell
' Date      : 15-Mar-2016
' Purpose   : Sub-routine of ValidateTrades and also called from PortfolioTradesToJuliaTrades.
'             Deals with the dates and flows that we pass in as semi-colon delimited strings
'---------------------------------------------------------------------------------------
Function ValidateFixedCashflows(TradeID As String, EndDate As Variant, Notional1 As Variant, Optional STK As SolumAddin.clsStacker)
          Dim DateArray As Variant
          Dim ErrString As String
          Dim FlowArray As Variant
          Dim i As Long
          Dim N As Long
          Dim ThrowOnError As Boolean
1         On Error GoTo ErrHandler
2         If STK Is Nothing Then ThrowOnError = True

3         If VarType(EndDate) <> vbString Then
4             ErrString = "For trade" + TradeID + " EndDate must be a semi-colon-delimited list of dates for example '15-May-2017;15-Jun-2017'"
5             If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
6         End If
7         If VarType(Notional1) <> vbString Then
8             ErrString = "For trade" + TradeID + " Notional1 must be a semi-colon-delimited list of amounts for example '1000000;500000'"
9             If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
10        End If
11        If VarType(EndDate) = vbString And VarType(Notional1) = vbString Then
12            DateArray = sTokeniseString(CStr(EndDate), ";")
13            FlowArray = sTokeniseString(CStr(Notional1), ";")
14            N = sNRows(DateArray)
15            If N <> sNRows(FlowArray) Then
16                ErrString = "For trade" + TradeID + " the number of dates listed in the 'End Date' column is not consistent with the number of cashflows listed in the 'Notional 1' column"
17                If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
18                Exit Function
19            End If
20            For i = 1 To N
21                If Not IsGoodDateString(CStr(DateArray(i, 1))) Then
22                    ErrString = "For trade" + TradeID + " the cashflow dates listed in the 'End Date' column are not in the correct format. Multiple dates are separated by semi-colons and must be in dd-mmm-yyyy format, for example: '15-May-2017;15-Jun-2017'."
23                    If Len(DateArray(i, 1)) < 20 Then ErrString = ErrString + " First date found in 'bad' format is: '" + CStr(DateArray(i, 1)) + "'"
24                    If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
25                    Exit Function
26                End If
27                If Not StringCanBeCastToDouble(CStr(FlowArray(i, 1))) Then
28                    ErrString = "For trade" + TradeID + " the cashflow amounts listed in the 'Notional 1' column are not in the correct format. An example of good format is: 1000000;500000"
29                    If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
30                    Exit Function
31                End If
32            Next i
33        End If
34        Exit Function
ErrHandler:
35        Throw "#ValidateFixedCashflows (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateFxStrip
' Author    : Philip Swannell
' Date      : 25 Dec 2016
' Purpose   : Sub-routine of ValidateTrades and also called from PortfolioTradesToJuliaTrades.
'             Deals with the dates and flows that we pass in as semi-colon delimited strings
'   If IsOption is TRUE then (elements of) Notional1 and Notional2 must be positive else must be of opposite sign
'---------------------------------------------------------------------------------------
Function ValidateFxStrip(TradeID As String, EndDate As Variant, Notional1 As Variant, Notional2 As Variant, IsOption As Boolean, Optional STK As SolumAddin.clsStacker)
          Dim DateArray As Variant
          Dim ErrString As String
          Dim FlowArray1 As Variant
          Dim FlowArray2 As Variant
          Dim i As Long
          Dim N As Long
          Dim ThrowOnError As Boolean
1         On Error GoTo ErrHandler
2         If STK Is Nothing Then ThrowOnError = True

3         If VarType(EndDate) <> vbString Then
4             ErrString = "For trade '" + TradeID + "' EndDate must be a semi-colon-delimited list of dates for example '15-May-2017;15-Jun-2017'"
5             If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
6         End If
7         If VarType(Notional1) <> vbString Then
8             ErrString = "For trade '" + TradeID + "' Notional1 must be a semi-colon-delimited list of amounts of Ccy 1, for example '1000000;500000'"
9             If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
10        End If
11        If VarType(Notional2) <> vbString Then
12            ErrString = "For trade '" + TradeID + "' Notional2 must be a semi-colon-delimited list of amounts of Ccy 2, for example '1000000;500000'"
13            If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
14        End If

15        If VarType(EndDate) = vbString And VarType(Notional1) = vbString And VarType(Notional2) = vbString Then
16            DateArray = sTokeniseString(CStr(EndDate), ";")
17            FlowArray1 = sTokeniseString(CStr(Notional1), ";")
18            FlowArray2 = sTokeniseString(CStr(Notional2), ";")
19            N = sNRows(DateArray)
20            If N <> sNRows(FlowArray1) Or N <> sNRows(FlowArray2) Then
21                ErrString = "For trade '" + TradeID + "' the number of dates in the 'End Date' column must equal the number of amounts in the 'Notional 1' column and the number of amounts in the 'Notional 2' column, but there are " + CStr(N) + ", " & CStr(sNRows(FlowArray1)) & " and " & CStr(sNRows(FlowArray2)) + " respectively."
22                If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
23                Exit Function
24            End If
25            For i = 1 To N
26                If Not IsGoodDateString(CStr(DateArray(i, 1))) Then
27                    ErrString = "For trade '" + TradeID + "' the dates listed in the 'End Date' column are not in the correct format. Multiple dates are separated by semi-colons and must be in dd-mmm-yyyy format, for example: '15-May-2017;15-Jun-2017' "
28                    If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
29                    Exit Function
30                End If
31                If Not StringCanBeCastToDouble(CStr(FlowArray1(i, 1))) Then
32                    ErrString = "For trade '" + TradeID + "' the amounts listed in the 'Notional 1' column are not in the correct format. An example of good format is: 1000000;500000"
33                    If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
34                    Exit Function
35                End If
36                If Not StringCanBeCastToDouble(CStr(FlowArray2(i, 1))) Then
37                    ErrString = "For trade '" + TradeID + "' the amounts listed in the 'Notional 2' column are not in the correct format. An example of good format is: 1000000;500000"
38                    If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
39                    Exit Function
40                End If

41                If IsOption Then
42                    If CDbl(FlowArray1(i, 1)) < 0 Then
43                        ErrString = "For trade '" + TradeID + "' the amounts listed in the 'Notional 1' column must be positive but element " + CStr(i) + " is negative"
44                        If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
45                        Exit Function
46                    ElseIf CDbl(FlowArray2(i, 1)) < 0 Then
47                        ErrString = "For trade '" + TradeID + "' the amounts listed in the 'Notional 2' column must be positive but element " + CStr(i) + " is negative"
48                        If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
49                        Exit Function
50                    End If
51                Else
52                    If Sgn(CDbl(FlowArray1(i, 1))) <> Sgn(CDbl(FlowArray2(i, 1))) Then
53                        ErrString = "For trade '" + TradeID + "' the amounts listed in the 'Notional 1' and 'Notional 2' columns must be of the same sign but that's not true for element " + CStr(i)
54                        If ThrowOnError Then Throw ErrString Else STK.StackData ErrString
55                        Exit Function
56                    End If
57                End If

58            Next i
59        End If
60        Exit Function
ErrHandler:
61        Throw "#ValidateFxStrip (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub TestFixedCashflowsDoubleClickHandler()
1         On Error GoTo ErrHandler
2         FixedCashflowsDoubleClickHandler ActiveCell, ActiveCell.Offset(0, 1), "Amount"
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestFixedCashflowsDoubleClickHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddressOfRange
' Author    : Philip Swannell
' Date      : 16-Mar-2016
' Purpose   : Returns the fully-qualified address of a range, which may have multiple areas.
'             designed to copy with space characters, exclamation marks etc in both the sheet
'             name and the workbook name. If origin range is passed and is on the same sheet or
'             in the same book as TheRange then the returned address is abreviated to exclude
'             the bookname or exclude both the book name and sheet name.
'             Feels like there should be an easier way to do all this! TheRange.Address(External:=True) is
'             nearly right but not quite what we want for multi-area ranges, since it returns
'             [BookName]SheetName!Area1Address,Area2Address
'             whereas we want:
'             [BookName]SheetName!Area1Address,[BookName]SheetName!Area2Address
'---------------------------------------------------------------------------------------
Function AddressOfRange(TheRange As Range, Optional OriginRange As Range)
          Dim addressFirstPart
          Dim CharactersToCheck As Variant
          Dim i As Long
          Dim includeBook As Boolean
          Dim includeSheet As Boolean
          Dim NeedEscapeCharacters As Boolean

1         On Error GoTo ErrHandler
2         If OriginRange Is Nothing Then
3             includeBook = True
4             includeSheet = True
5         Else
6             If TheRange.Parent Is OriginRange.Parent Then
7                 includeBook = False
8                 includeSheet = False
9             ElseIf TheRange.Parent.Parent Is OriginRange.Parent.Parent Then
10                includeBook = False
11                includeSheet = True
12            Else
13                includeBook = True
14                includeSheet = True
15            End If
16        End If

          'https://d.docs.live.net/4251b448d4115355/Excel Sheets/[Find bad characters in sheet names.xlsm]Sheet3
17        CharactersToCheck = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, _
              19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, _
              35, 36, 37, 38, 39, 40, 41, 43, 44, 45, 59, 60, 61, 62, 64, 94, _
              96, 123, 124, 125, 126, 127, 129, 130, 132, 139, 141, 143, 144, _
              149, 155, 157, 160, 162, 163, 165, 166, 169, 171, 172, 174, 187)

18        If includeBook Then
19            For i = LBound(CharactersToCheck) To UBound(CharactersToCheck)
20                If InStr(TheRange.Parent.Parent.Name, Chr(CharactersToCheck(i))) > 0 Then
21                    NeedEscapeCharacters = True
22                    Exit For
23                End If
24            Next i
25        End If
26        If includeSheet Then
27            If Not NeedEscapeCharacters Then
28                For i = LBound(CharactersToCheck) To UBound(CharactersToCheck)
29                    If InStr(TheRange.Parent.Name, Chr(CharactersToCheck(i))) > 0 Then
30                        NeedEscapeCharacters = True
31                        Exit For
32                    End If
33                Next i
34            End If
35        End If

36        If Not includeSheet And Not includeBook Then
37            addressFirstPart = ""
38        ElseIf includeSheet And Not includeBook Then
39            If NeedEscapeCharacters Then
40                addressFirstPart = "'" + Replace(TheRange.Parent.Name, "'", "''") + "'!"
41            Else
42                addressFirstPart = TheRange.Parent.Name + "!"
43            End If
44        Else
45            If NeedEscapeCharacters Then
46                addressFirstPart = "'[" + Replace(TheRange.Parent.Parent.Name + "]" + TheRange.Parent.Name, "'", "''") + "'!"
47            Else
48                addressFirstPart = "[" + TheRange.Parent.Parent.Name + "]" + TheRange.Parent.Name + "!"
49            End If
50        End If

51        If InStr(TheRange.Address, ",") = 0 Or addressFirstPart = "" Then
52            AddressOfRange = addressFirstPart + TheRange.Address
53        Else
54            AddressOfRange = sConcatenateStrings(sArrayConcatenate(addressFirstPart, sTokeniseString(TheRange.Address, ",")), ",")
55        End If
56        Exit Function
ErrHandler:
57        Throw "#AddressOfRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : NotionalDoubleClickHandler
' Author    : Philip
' Date      : 09-Apr-2016
' Purpose   : Double-click handler for notionals for interest rate swaps
'---------------------------------------------------------------------------------------
Function NotionalDoubleClickHandler(NotionalCell As Range)
          Dim c As Range
          Dim CopiedRange As Range
          Dim DefaultValue As String
          Dim Prompt As String
          Dim resRange As Range
          Dim Title As String

          Const BadRangesPrompt = "Selected range must have one column."
          Const BadNotionalsSelected = "Cells selected must contain positive numbers."

1         On Error GoTo ErrHandler
2         Prompt = "Select a column of cells containing the notionals" + vbLf + "(one per period) for this leg of the trade." + vbLf + "Tip: Try copying the cells first (Ctrl C)."
3         Title = "Amortising Notional"

4         On Error Resume Next
5         GetCopiedRange CopiedRange
6         On Error GoTo ErrHandler
7         If Not CopiedRange Is Nothing Then
8             DefaultValue = AddressOfRange(CopiedRange, NotionalCell)
9         End If

TryAgain:
10        Set resRange = Nothing
11        On Error Resume Next
12        Set resRange = Range(InputBoxPlus(Prompt, Title, DefaultValue, , , , , , , , , , True))

13        On Error GoTo ErrHandler
14        If Not CopiedRange Is Nothing Then
15            If Application.CutCopyMode = False Then
16                CopiedRange.Copy
17            End If
18        End If

19        If resRange Is Nothing Then
20            Exit Function
21        End If
22        If resRange.Areas.Count > 1 Then
23            MsgBoxPlus "Multiple selections are not allowed.", vbInformation, Title
24            GoTo TryAgain
25        ElseIf resRange.Columns.Count > 1 Then
26            MsgBoxPlus BadRangesPrompt, vbInformation, Title
27            GoTo TryAgain
28        End If

29        For Each c In resRange.Cells
30            If Not IsNumberOrDate(c.Value) Then
31                MsgBoxPlus BadNotionalsSelected, vbInformation, Title
32                GoTo TryAgain
33            End If
34        Next c
35        If resRange.Cells.Count = 1 Then
36            NotionalDoubleClickHandler = resRange.Value
37        Else
              Dim NotionalsString As String
38            For Each c In resRange.Cells
39                NotionalsString = NotionalsString + ";" + CStr(c.Value2)
40            Next c
41            NotionalsString = Mid(NotionalsString, 2)
42            NotionalDoubleClickHandler = NotionalsString
43        End If

44        Exit Function
ErrHandler:
45        Throw "#NotionalDoubleClickHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FixedCashflowsDoubleClickHandler
' Author    : Philip Swannell
' Date      : 16-Mar-2016
' Purpose   : Double-click handler to make it easy to enter dates and amounts for trades of type FixedCashflows.
'---------------------------------------------------------------------------------------
Function FixedCashflowsDoubleClickHandler(DatesCell As Range, AmountsCell As Range, Mode As String)
          Dim AmountsRange As Range
          Dim CopiedRange As Range
          Dim CopyOfErr As String
          Dim DatesRange As Range
          Dim DefaultValue As String
          Dim OldBCE As Boolean
          Dim Prompt As String
          Dim resRange As Range
          Dim Title As String
          Const BadRangesPrompt = "Selected range must have one or two columns, or you can select two one column ranges, one for dates, one for amounts."
          Const BadDatesSelected = "Some of the cells selected for dates do not contain dates"
          Const BadAmountsSelected = "Some of the cells selected for amounts do not contain amounts"
          Dim SUH As SolumAddin.clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         If LCase(Mode) = "amount" Then
4             Prompt = "Select a column of amounts, or two columns of dates and amounts." + vbLf + "Tip: Try copying the range (Ctrl C) first."
5             Title = "FixedCashflows Amounts"
6         Else
7             Prompt = "Select a column of dates, or two columns of dates and amounts." + vbLf + "Tip: Try copying the range (Ctrl C) first."
8             Title = "FixedCashflows Dates"
9         End If
10        On Error Resume Next
11        GetCopiedRange CopiedRange
12        On Error GoTo ErrHandler
13        If Not CopiedRange Is Nothing Then
14            DefaultValue = AddressOfRange(CopiedRange, DatesCell)
15        End If

TryAgain:
16        Set resRange = Nothing
17        On Error Resume Next
          ' Set resRange = Application.InputBox(Prompt, Title, DefaultValue, , , , , 8)
18        Set resRange = Range(InputBoxPlus(Prompt, Title, DefaultValue, , , 300, , , , , , , True))
19        On Error GoTo ErrHandler

20        If resRange Is Nothing Then
21            Exit Function
22        End If
23        Set SUH = CreateScreenUpdateHandler()

24        If resRange.Areas.Count > 2 Or resRange.Columns.Count > 2 Then
25            MsgBoxPlus BadRangesPrompt, vbInformation, Title
26            GoTo TryAgain
27        End If

28        If resRange.Areas.Count = 2 Then
29            If resRange.Areas(1).Columns.Count <> 1 Or _
                  resRange.Areas(2).Columns.Count <> 1 Or _
                  resRange.Areas(1).Rows.Count <> resRange.Areas(2).Rows.Count Or _
                  Not (Application.Intersect(resRange.Areas(1), resRange.Areas(2)) Is Nothing) Then
30                MsgBoxPlus BadRangesPrompt, vbInformation, Title
31                GoTo TryAgain
32            End If
33        End If

34        If resRange.Areas.Count = 2 Then
35            Set DatesRange = resRange.Areas(1)
36            Set AmountsRange = resRange.Areas(2)
37        ElseIf resRange.Columns.Count = 1 Then
38            If LCase(Mode) = "amount" Then
39                Set AmountsRange = resRange
40            Else
41                Set DatesRange = resRange
42            End If
43        ElseIf resRange.Columns.Count = 2 Then
44            Set DatesRange = resRange.Columns(1)
45            Set AmountsRange = resRange.Columns(2)
46        End If

          Dim AmountsString As String
          Dim c As Range
          Dim DatesString As String

47        If Not DatesRange Is Nothing Then
48            For Each c In DatesRange.Cells
49                If Not IsNumberOrDate(c.Value) Then
50                    MsgBoxPlus BadDatesSelected, vbInformation, Title
51                    GoTo TryAgain
52                ElseIf c.Value < 0 Then
53                    MsgBoxPlus BadDatesSelected, vbInformation, Title
54                    GoTo TryAgain

55                ElseIf CLng(c.Value <> c.Value) Then
56                    MsgBoxPlus BadDatesSelected, vbInformation, Title
57                    GoTo TryAgain
58                End If
59            Next c
60        End If

61        If Not AmountsRange Is Nothing Then
62            For Each c In AmountsRange.Cells
63                If Not IsNumberOrDate(c.Value) Then
64                    MsgBoxPlus BadAmountsSelected, vbInformation, Title
65                    GoTo TryAgain
66                End If
67            Next c
68        End If

69        If Not DatesRange Is Nothing Then
70            For Each c In DatesRange.Cells
71                DatesString = DatesString + ";" + FormatAsDate(c.Value2, "d-mmm-yyyy")
72            Next c
73            DatesString = Mid(DatesString, 2)
74        End If

75        If Not AmountsRange Is Nothing Then
76            For Each c In AmountsRange.Cells
77                AmountsString = AmountsString + ";" + CStr(c.Value2)
78            Next c
79            AmountsString = Mid(AmountsString, 2)
80        End If

          Dim DoingBoth As Boolean
81        DoingBoth = (Not DatesRange Is Nothing) And (Not AmountsRange Is Nothing)

82        If DoingBoth Then
83            BackUpRange Application.Union(DatesCell, AmountsCell), shUndo, , True
84            gBlockChangeEvent = True
85            DatesCell.Value = "'" + DatesString
86            gBlockChangeEvent = OldBCE
87            AmountsCell.Value = "'" + AmountsString
88            FormatTradesRange DatesCell
89            Application.OnUndo "Restore " + Replace(DatesCell.Address, "$", "") + " and " & Replace(AmountsCell.Address, "$", "") + " to their previous values", "RestoreRange"
90        ElseIf Not DatesRange Is Nothing Then
91            BackUpRange DatesCell, shUndo, , True
92            DatesCell.Value = "'" + DatesString
93            FormatTradesRange DatesCell
94            Application.OnUndo "Restore " + Replace(DatesCell.Address, "$", "") + " to its previous value", "RestoreRange"
95        ElseIf Not AmountsRange Is Nothing Then
96            BackUpRange AmountsCell, shUndo, , True
97            AmountsCell.Value = "'" + AmountsString
98            FormatTradesRange AmountsCell
99            Application.OnUndo "Restore " + Replace(AmountsCell.Address, "$", "") + " to its previous value", "RestoreRange"
100       End If

101       Exit Function
ErrHandler:
102       CopyOfErr = "#FixedCashflowsDoubleClickHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
103       gBlockChangeEvent = OldBCE
104       Throw CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : FormatAsDate
' Author    : Philip Swannell
' Date      : 01-Jan-2017
' Purpose   : Avoid uninformative error "Overflow"
'---------------------------------------------------------------------------------------
Function FormatAsDate(TheValue, TheFormat As String)
1         On Error GoTo ErrHandler
2         FormatAsDate = Format(TheValue, TheFormat)
3         Exit Function
ErrHandler:
4         Throw "Cannot express the value '" + CStr(TheValue) + "' as a date"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FxStripDoubleClickHandler
' Author    : Philip Swannell
' Date      : 16-Mar-2016
' Purpose   : Double-click handler to make it easy to enter dates and amounts for trades
'             of type FxOptionStrip and FxForwardStrip.
'---------------------------------------------------------------------------------------
Sub FxStripDoubleClickHandler(ValuationFunction As String, DatesCell As Range, Amounts1Cell As Range, Amounts2Cell As Range, Mode As String)
          Dim Amounts1Range As Range
          Dim Amounts2Range As Range
          Dim CopiedRange As Range
          Dim DatesRange As Range
          Dim DefaultValue As String
          Dim Prompt As String
          Dim resRange As Range
          Dim Title As String
          Const BadRangesPrompt = "Selected range must have one or three columns, or you can select three one column ranges - one for dates and two for amounts."
          Const BadDatesSelected = "Some of the cells selected for dates do not contain dates"
          Const BadAmountsSelected = "Some of the cells selected for amounts do not contain amounts"
          Dim CopyOfErr As String
          Dim OldBCE As Boolean
          Dim SUH As SolumAddin.clsScreenUpdateHandler

1         On Error GoTo ErrHandler

2         OldBCE = gBlockChangeEvent

3         If LCase(Mode) = "amount" Then
4             Prompt = "Select a column of amounts, or three columns - one for dates and" + vbLf + "two for amounts." + vbLf + "Tip: Try copying the range (Ctrl C) first."
5             Title = ValuationFunction + " Amounts"
6         Else
7             Prompt = "Select a column of dates, or three columns - one for dates" + vbLf + "and two for amounts." + vbLf + "Tip: Try copying the range (Ctrl C) first."
8             Title = ValuationFunction + " Dates"
9         End If
10        On Error Resume Next
11        GetCopiedRange CopiedRange
12        On Error GoTo ErrHandler
13        If Not CopiedRange Is Nothing Then
14            DefaultValue = AddressOfRange(CopiedRange, DatesCell)
15        End If

TryAgain:
16        Set resRange = Nothing
17        On Error Resume Next
18        Set resRange = Range(InputBoxPlus(Prompt, Title, DefaultValue, , , 300, , , , , , , True))
19        On Error GoTo ErrHandler

20        If resRange Is Nothing Then
21            Exit Sub
22        End If
23        Set SUH = CreateScreenUpdateHandler()

          Dim a As Range
          Dim NumRows As Long
          Dim TotalCols As Long
24        NumRows = resRange.Rows.Count
25        For Each a In resRange.Areas
26            If a.Rows.Count <> NumRows Then
27                MsgBoxPlus BadRangesPrompt, vbInformation, Title
28                GoTo TryAgain
29            End If
30            TotalCols = TotalCols + a.Columns.Count
31        Next a

32        If TotalCols <> 1 And TotalCols <> 3 Then
33            MsgBoxPlus BadRangesPrompt, vbInformation, Title
34            GoTo TryAgain
35        End If
36        If TotalCols = 3 Then
37            Set DatesRange = resRange.Areas(1).Columns(1)
38            If resRange.Areas(1).Columns.Count > 1 Then
39                Set Amounts1Range = resRange.Areas(1).Columns(2)
40            Else
41                Set Amounts1Range = resRange.Areas(2).Columns(1)
42            End If
43            If resRange.Areas(1).Columns.Count > 2 Then
44                Set Amounts2Range = resRange.Areas(1).Columns(3)
45            ElseIf resRange.Areas(1).Columns.Count = 2 Then
46                Set Amounts2Range = resRange.Areas(2).Columns(1)
47            ElseIf resRange.Areas(1).Columns.Count = 1 And resRange.Areas(2).Columns.Count = 1 Then
48                Set Amounts2Range = resRange.Areas(3).Columns(1)
49            End If
50        ElseIf resRange.Columns.Count = 1 Then
51            If LCase(Mode) = "amount" Then
52                Set Amounts1Range = resRange
53            Else
54                Set DatesRange = resRange
55            End If
56        End If

          Dim Amounts1String As String
          Dim Amounts2String As String
          Dim c As Range
          Dim DatesString As String

57        If Not DatesRange Is Nothing Then
58            For Each c In DatesRange.Cells
59                If Not IsNumberOrDate(c.Value) Then
60                    MsgBoxPlus BadDatesSelected, vbInformation, Title
61                    GoTo TryAgain
62                ElseIf c.Value < 0 Then
63                    MsgBoxPlus BadDatesSelected, vbInformation, Title
64                    GoTo TryAgain

65                ElseIf CLng(c.Value <> c.Value) Then
66                    MsgBoxPlus BadDatesSelected, vbInformation, Title
67                    GoTo TryAgain
68                End If
69            Next c
70        End If

71        If Not Amounts1Range Is Nothing Then
72            For Each c In Amounts1Range.Cells
73                If Not IsNumberOrDate(c.Value) Then
74                    MsgBoxPlus BadAmountsSelected, vbInformation, Title
75                    GoTo TryAgain
76                End If
77            Next c
78        End If

79        If Not DatesRange Is Nothing Then
80            For Each c In DatesRange.Cells
81                DatesString = DatesString + ";" + FormatAsDate(c.Value2, "d-mmm-yyyy")
82            Next c
83            DatesString = Mid(DatesString, 2)
84        End If

85        If Not Amounts1Range Is Nothing Then
86            For Each c In Amounts1Range.Cells
87                Amounts1String = Amounts1String + ";" + CStr(c.Value2)
88            Next c
89            Amounts1String = Mid(Amounts1String, 2)
90        End If

91        If Not Amounts2Range Is Nothing Then
92            For Each c In Amounts2Range.Cells
93                Amounts2String = Amounts2String + ";" + CStr(c.Value2)
94            Next c
95            Amounts2String = Mid(Amounts2String, 2)
96        End If

97        If TotalCols = 3 Then
98            BackUpRange Application.Union(DatesCell, Amounts1Cell, Amounts2Cell), shUndo, , True
99            gBlockChangeEvent = True
100           DatesCell.Value = "'" + DatesString
101           Amounts1Cell.Value = "'" + Amounts1String
102           gBlockChangeEvent = OldBCE
103           Amounts2Cell.Value = "'" + Amounts2String
104           Application.OnUndo "Restore cells " + Replace(Application.Union(DatesCell, Amounts1Cell, Amounts2Cell).Address, "$", "") + " to their previous values", "RestoreRange"
105       ElseIf LCase(Mode) = "amount" Then
106           BackUpRange ActiveCell, shUndo, , True
107           ActiveCell.Value = "'" + Amounts1String
108           Application.OnUndo "Restore cell " + Replace(ActiveCell.Address, "$", "") + " to its previous value", "RestoreRange"
109       Else
110           BackUpRange ActiveCell, shUndo, , True
111           ActiveCell.Value = "'" + DatesString
112           Application.OnUndo "Restore cell " + Replace(ActiveCell.Address, "$", "") + " to its previous value", "RestoreRange"
113       End If

114       Exit Sub
ErrHandler:
115       CopyOfErr = "#FxStripDoubleClickHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
116       gBlockChangeEvent = OldBCE
117       Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsValidNotional
' Author    : Philip
' Date      : 09-Apr-2016
' Purpose   : Tests for non-negative number or semi-colon delimited list of non-negative numbers
'---------------------------------------------------------------------------------------
Function IsValidNotional(TestThis As Variant)
1         On Error GoTo ErrHandler
2         If Not gDoValidation Then
3             IsValidNotional = True
4             Exit Function
5         End If
6         If IsNumber(TestThis) Then
7             IsValidNotional = TestThis >= 0
8         ElseIf VarType(TestThis) = vbString Then
9             If InStr(TestThis, ";") = 0 Then
10                IsValidNotional = False
11                Exit Function
12            End If
              Dim i As Long
              Dim Res As Variant
13            Res = sTokeniseString(CStr(TestThis), ";")
14            For i = 1 To sNRows(Res)
15                If Not StringCanBeCastToDouble(CStr(Res(i, 1))) Then
16                    IsValidNotional = False
17                    Exit Function
18                ElseIf CDbl(Res(i, 1)) < 0 Then
19                    IsValidNotional = False
20                    Exit Function
21                End If
22            Next i
23            IsValidNotional = True
24        Else
25            IsValidNotional = False
26        End If

27        Exit Function
ErrHandler:
28        Throw "#IsValidNotional (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FlipNotionalSign
' Author    : Philip
' Date      : 09-Apr-2016
' Purpose   : Given a semi-colon-delimited list of notionals, returns a semi-colon delimited list of minus those notionals
'---------------------------------------------------------------------------------------
Function FlipNotionalSign(Notionals As String)
          Dim i As Long
          Dim LB As Long
          Dim NotionalsArray
          Dim Result As String
          Dim UB As Long

1         On Error GoTo ErrHandler
2         NotionalsArray = VBA.Split(Notionals, ";")
3         LB = LBound(NotionalsArray)
4         UB = UBound(NotionalsArray)

5         Result = CStr(-CDbl(NotionalsArray(LB)))

6         For i = LB + 1 To UB
7             Result = Result + ";" + CStr(-CDbl(NotionalsArray(i)))
8         Next
9         FlipNotionalSign = Result
10        Exit Function
ErrHandler:
11        Throw "#FlipNotionalSign (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
