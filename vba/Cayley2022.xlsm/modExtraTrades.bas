Attribute VB_Name = "modExtraTrades"
' -----------------------------------------------------------------------------------------------------------------------
' Name: modExtraTrades
' Kind: Module
' Purpose: Create the "Additional trades" in the format that Julia understands. These trades are described in abbreviated
'          format at the top of the CreditUsageSheet and are the "unit of account" for calculating trade headroom.
'          Note that trades are represented from the BANKS' POINT OF VIEW!
' Author: Philip Swannell
' Date: 18-Feb-2022
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowExtraTrades
' Author     : Philip Swannell
' Date       : 21-Feb-2022
' Purpose    : For development work, make it possible to view the "ExtraTrades"
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowExtraTrades()
          Dim ExtraTrades As Variant

1         On Error GoTo ErrHandler
2         If gModel_CM Is Nothing Then
3             OpenOtherBooks
4             BuildModelsInJulia True, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")
5         End If

6         ExtraTrades = ConstructExtraTrades(RangeFromSheet(shCreditUsage, "ExtraTradesAre"), _
              gModel_CM, _
              RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value, _
              RangeFromSheet(shCreditUsage, "ExtraTradeLabels").Value, _
              True)

7         g ExtraTrades

8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#ShowExtraTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub ParseExtraTradesAre(ExtraTradesAre As String, ByRef FxIncluded As Boolean, RatesIncluded As Boolean)
1         On Error GoTo ErrHandler
2         If Left(ExtraTradesAre, 2) = "Fx" Then
3             FxIncluded = True
4             RatesIncluded = False
5         ElseIf Left(ExtraTradesAre, 3) = "IRS" Then
6             FxIncluded = False
7             RatesIncluded = True
8         Else
9             Throw "ExtraTradesAre of '" & ExtraTradesAre & "' is not recognised. Valid values are:" _
                  & vbLf & sConcatenateStrings(AllowedExtraTradesAre(False), vbLf)
10        End If

11        Exit Sub
ErrHandler:
12        Throw "#ParseExtraTradesAre (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AllowedExtraTradesAre
' Author     : Philip Swannell
' Date       : 18-Feb-2022
' Purpose    : Returns a list (2d, 1col array) of allowed values for the variable ExtraTradesAre
' Parameters : If True the return is shorter, listing only likely-to-be-useful values
' -----------------------------------------------------------------------------------------------------------------------
Function AllowedExtraTradesAre(MajorCurrenciesOnly As Boolean)
          Dim CCyArray
          Static OldCcys As String
          Static OldMCO As Boolean
          Static Res
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim MatchRes
          Dim N As Long
          Const DefaultOption = "Fx Airbus sells USD, buys EUR"
          Const MajorCurrencies = "EUR,USD"

1         On Error GoTo ErrHandler

2         If OldCcys <> RangeFromSheet(shConfig, "CurrenciesToInclude") Or MajorCurrenciesOnly <> OldMCO Then
3             OldCcys = RangeFromSheet(shConfig, "CurrenciesToInclude")
4             OldMCO = MajorCurrenciesOnly
5             CCyArray = sSortedArray(sTokeniseString(RangeFromSheet(shConfig, "CurrenciesToInclude")))
6             If MajorCurrenciesOnly Then
7                 CCyArray = sCompareTwoArrays(CCyArray, sTokeniseString(MajorCurrencies), "Common,NoHeaders")
8             End If

9             N = sNRows(CCyArray)

10            Res = sReshape("", N * (N - 1), 1)
11            For i = 1 To N
12                For j = 1 To N
13                    If i <> j Then
14                        k = k + 1
15                        Res(k, 1) = "Fx Airbus sells " & CCyArray(i, 1) & ", buys " & CCyArray(j, 1)
16                    End If
17                Next j
18            Next i

19            Res = sArrayStack(Res, _
                  "FxOption Airbus sells ATM call on EUR vs USD", _
                  "FxOption Airbus sells ATM put on EUR vs USD", _
                  sArrayConcatenate("IRS Airbus receives fixed ", CCyArray), _
                  sArrayConcatenate("IRS Airbus pays fixed ", CCyArray))
                  
              'Make the default option be at the top
20            MatchRes = sMatch(DefaultOption, Res)
21            If IsNumber(MatchRes) Then
22                If MatchRes <> 1 Then
23                    Res = sArrayStack(DefaultOption, sCompareTwoArrays(DefaultOption, Res, "In2AndNotIn1,NoHeaders"))
24                End If
25            End If
                  
26        End If
27        AllowedExtraTradesAre = Res

28        Exit Function
ErrHandler:
29        Throw "#AllowedExtraTradesAre (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConstructExtraTrades
' Author     : Philip Swannell
' Date       : 18-Feb-2022
' Purpose    :
' Parameters :
'  ExtraTradesAre   :
'  ModelBareBones   :
'  Amounts          :
'  Labels           :
'  OnlyNonZeroTrades:
' -----------------------------------------------------------------------------------------------------------------------
Function ConstructExtraTrades(ExtraTradesAre As String, ModelBareBones As Dictionary, _
          Amounts, Labels, Optional OnlyNonZeroTrades As Boolean = True)
                
1         On Error GoTo ErrHandler

          Const LabelsError = "Labels must be strings of the form NY for some integer N"
          Dim AnchorDate As Date
          Dim i As Long
          Dim MaturityDates() As Date
          Dim N As Long
          Dim NY As String
          
2         If Not IsNumber(sMatch(ExtraTradesAre, AllowedExtraTradesAre(True))) Then
3             Throw "'" & ExtraTradesAre & "' is not valid input to the 'ExtraTradesAre` cell on the CreditUsage" & _
                  " worksheet. Valid values are:" + vbLf + sConcatenateStrings(AllowedExtraTradesAre(True), vbLf)
4         End If

5         AnchorDate = DictGet(ModelBareBones, "AnchorDate")
6         Force2DArrayRMulti Amounts, Labels

7         N = sNRows(Amounts)
8         ReDim MaturityDates(1 To N, 1 To 1)

9         If sNCols(Amounts) > 1 Then Throw "Amounts must be a single column array"
10        If sNCols(Labels) > 1 Then Throw "Labels must be a single column array"
11        If sNRows(Amounts) <> sNRows(Labels) Then Throw "Amounts and Labels must have the same number of rows"

12        For i = 1 To N
13            If VarType(Labels(i, 1)) <> vbString Then Throw LabelsError
14            If Right(CStr(Labels(i, 1)), 1) <> "Y" Then Throw LabelsError
15            NY = Left(Labels(i, 1), Len(Labels(i, 1)) - 1)
16            If Not IsNumeric(NY) Then Throw LabelsError
17            NY = CDbl(NY)
18            MaturityDates(i, 1) = AnchorDate + 365 * NY
19        Next i
20        For i = 1 To N
21            If Not IsNumber(Amounts(i, 1)) Then
22                Throw "Amounts must be numbers, but element " & CStr(i) & " is '" & CStr(Amounts(i, 1)) & "'"
23            End If
24        Next i

25        If ExtraTradesAre = "Fx Airbus sells USD, buys EUR" Then
26            ConstructExtraTrades = ConstructExtraFxForwards(ModelBareBones, "USD", "EUR", Amounts, _
                  MaturityDates, OnlyNonZeroTrades)
27        ElseIf ExtraTradesAre = "Fx Airbus sells EUR, buys USD" Then
28            ConstructExtraTrades = ConstructExtraFxForwards(ModelBareBones, "USD", "EUR", sArrayMultiply(Amounts, -1), _
                  MaturityDates, OnlyNonZeroTrades)
29        ElseIf Left$(ExtraTradesAre, 26) = "IRS Airbus receives fixed " Then
30            ConstructExtraTrades = ConstructExtraInterestRateSwaps(ModelBareBones, UCase(Right$(ExtraTradesAre, 3)), _
                  sArrayMultiply(Amounts, -1), MaturityDates, OnlyNonZeroTrades)
31        ElseIf Left$(ExtraTradesAre, 22) = "IRS Airbus pays fixed " Then
32            ConstructExtraTrades = ConstructExtraInterestRateSwaps(ModelBareBones, UCase(Right$(ExtraTradesAre, 3)), _
                  Amounts, MaturityDates, OnlyNonZeroTrades)
33        ElseIf ExtraTradesAre = "FxOption Airbus sells ATM call on EUR vs USD" Then
34            ConstructExtraTrades = ConstructExtraFxOptions(ModelBareBones, "USD", _
                  Amounts, MaturityDates, False, OnlyNonZeroTrades)
35        ElseIf ExtraTradesAre = "FxOption Airbus sells ATM put on EUR vs USD" Then
36            ConstructExtraTrades = ConstructExtraFxOptions(ModelBareBones, "USD", _
                  Amounts, MaturityDates, True, OnlyNonZeroTrades)
37        Else
38            Throw "String '" + ExtraTradesAre + "' is not recognised as valid input for ExtraTradesAre"
39        End If

40        Exit Function
ErrHandler:
41        Throw "#ConstructExtraTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConstructExtraInterestRateSwaps
' Author     : Philip Swannell
' Date       : 18-Feb-2022
' Purpose    :
' Parameters :
'  ModelBareBones   :
'  CCY              :
'  Amounts          :
'  MaturityDates    :
'  OnlyNonZeroTrades:
' -----------------------------------------------------------------------------------------------------------------------
Function ConstructExtraInterestRateSwaps(ModelBareBones As Dictionary, Ccy As String, Amounts, MaturityDates() As Date, _
          OnlyNonZeroTrades As Boolean)

          Dim AnchorDate As Date
          Dim i As Long
          Dim k As Long
          Dim NTrades As Long
          Dim SwapRates As Dictionary
          Dim Trades() As Variant
          Const cnTradeID As Long = 1
          Const cnValuationFunction As Long = 2
          Const cnCounterparty As Long = 3
          Const cnStartDate As Long = 4
          Const cnEndDate As Long = 5
          Const cnReceiveCurrency As Long = 6
          Const cnPayCurrency As Long = 7
          Const cnReceiveNotional As Long = 8
          Const cnPayNotional As Long = 9
          Const cnPayBDC As Long = 10
          Const cnPayCoupon As Long = 11
          Const cnPayDCT As Long = 12
          Const cnPayFrequency As Long = 13
          Const cnPayIndex As Long = 14
          Const cnReceiveBDC As Long = 15
          Const cnReceiveCoupon As Long = 16
          Const cnReceiveDCT As Long = 17
          Const cnReceiveFrequency As Long = 18
          Const cnReceiveIndex As Long = 19

1         On Error GoTo ErrHandler

2         If OnlyNonZeroTrades Then
3             For i = 1 To sNRows(Amounts)
4                 If Amounts(i, 1) <> 0 Then
5                     NTrades = NTrades + 1
6                 End If
7             Next i
8         Else
9             NTrades = sNRows(Amounts)
10        End If
11        AnchorDate = DictGet(ModelBareBones, "AnchorDate")
12        Set SwapRates = GetItem(gMarketData, "SwapRates_" & Ccy)
          
13        ReDim Trades(1 To NTrades + 1, 1 To 19)
14        Trades(1, cnTradeID) = "TradeID"
15        Trades(1, cnValuationFunction) = "ValuationFunction"
16        Trades(1, cnCounterparty) = "Counterparty"
17        Trades(1, cnStartDate) = "StartDate"
18        Trades(1, cnEndDate) = "EndDate"
19        Trades(1, cnReceiveCurrency) = "ReceiveCurrency"
20        Trades(1, cnPayCurrency) = "PayCurrency"
21        Trades(1, cnReceiveNotional) = "ReceiveNotional"
22        Trades(1, cnPayNotional) = "PayNotional"
23        Trades(1, cnPayBDC) = "PayBDC"
24        Trades(1, cnPayCoupon) = "PayCoupon"
25        Trades(1, cnPayDCT) = "PayDCT"
26        Trades(1, cnPayFrequency) = "PayFrequency"
27        Trades(1, cnPayIndex) = "PayIndex"
28        Trades(1, cnReceiveBDC) = "ReceiveBDC"
29        Trades(1, cnReceiveCoupon) = "ReceiveCoupon"
30        Trades(1, cnReceiveDCT) = "ReceiveDCT"
31        Trades(1, cnReceiveFrequency) = "ReceiveFrequency"
32        Trades(1, cnReceiveIndex) = "ReceiveIndex"
          
          'Construct an on-market trade. Would be better to solve for the fixed rate on the Julia side
33        k = 1
34        For i = 1 To sNRows(Amounts)
35            If Amounts(i, 1) <> 0 Or Not OnlyNonZeroTrades Then
36                k = k + 1
37                Trades(k, cnTradeID) = "ExtraTrade" & CStr(k - 1)
38                Trades(k, cnValuationFunction) = "InterestRateSwap"
39                Trades(k, cnCounterparty) = "Not Specified"
40                Trades(k, cnStartDate) = AnchorDate
41                Trades(k, cnEndDate) = MaturityDates(i, 1)
42                Trades(k, cnReceiveCurrency) = Ccy
43                Trades(k, cnPayCurrency) = Ccy
44                Trades(k, cnReceiveNotional) = Amounts(i, 1)
45                Trades(k, cnPayNotional) = -Amounts(i, 1)
46                Trades(k, cnPayBDC) = "Mod Foll"
47                Trades(k, cnPayCoupon) = 0 ' i.e. zero margin
48                Trades(k, cnPayDCT) = LastElementOf(GetItem(SwapRates, "FloatingDCT"))
49                Trades(k, cnPayFrequency) = LastElementOf(GetItem(SwapRates, "FloatingFrequency"))
50                Trades(k, cnPayIndex) = "Libor" 'TODO change to OIS?
51                Trades(k, cnReceiveBDC) = "Mod Foll"
52                Trades(k, cnReceiveCoupon) = SwapRateFromMarketData(Ccy, CLng((MaturityDates(i, 1) - AnchorDate) / 365.25))
53                Trades(k, cnReceiveDCT) = LastElementOf(GetItem(SwapRates, "FixedDCT"))
54                Trades(k, cnReceiveFrequency) = LastElementOf(GetItem(SwapRates, "FixedFrequency"))
55                Trades(k, cnReceiveIndex) = "Fixed"
56            End If
57        Next i

58        ConstructExtraInterestRateSwaps = Trades

59        Exit Function
ErrHandler:
60        Throw "#ConstructExtraInterestRateSwaps (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub testSwapRateFromMarketData()
1         On Error GoTo ErrHandler
2         If gMarketData Is Nothing Then BuildModelsInJulia True, 1, 1

3         Debug.Print SwapRateFromMarketData("EUR", 8)
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#testSwapRateFromMarketData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function SwapRateFromMarketData(Ccy As String, NYears As Long)

          Dim i As Long
          Dim InterpRes
          Dim Rates As Collection
          Dim SwapRates As Dictionary
          Dim Tenors As Collection

1         On Error GoTo ErrHandler
2         If gMarketData Is Nothing Then Throw "Cannot find Dictionary 'gMarketData'"

3         Set SwapRates = gMarketData("SwapRates_" & UCase(Ccy))
4         Set Tenors = SwapRates("Tenors")
5         Set Rates = SwapRates("Rates")
          'Hope for exact match
6         For i = 1 To Tenors.Count
7             If Tenors(i) = CStr(NYears) & "Y" Then
8                 SwapRateFromMarketData = Rates(i)
9                 Exit Function
10            End If
11        Next i
          
          'Otherwise interpolate
          Dim RatesVec
          Dim TenorsVec

12        TenorsVec = CollectionToColumn(Tenors)
13        RatesVec = CollectionToColumn(Rates)
          'Convert from strings to numbers
14        For i = 1 To UBound(TenorsVec)
15            TenorsVec(i, 1) = TenorToTimeCore(TenorsVec(i, 1))
16        Next i

          'Interpolate with flat extrapolation at either end
17        InterpRes = ThrowIfError(sInterp(TenorsVec, RatesVec, NYears, , "FF"))
18        SwapRateFromMarketData = InterpRes(1, 1)

19        Exit Function
ErrHandler:
20        Throw "#SwapRateFromMarketData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CollectionToColumn
' Author     : Philip Swannell
' Date       : 21-Feb-2022
' Purpose    : Painful, the gMarketData dictionary contains collections where IMO it should contain vectors, use this
'              method to convert to 2-dim array with one column (which is what SolumAddin functions generally expect).
' -----------------------------------------------------------------------------------------------------------------------
Function CollectionToColumn(c As Collection)
          Dim i As Long

          Dim Res() As Variant
1         On Error GoTo ErrHandler
2         ReDim Res(1 To c.Count, 1 To 1)

3         For i = 1 To c.Count
4             Res(i, 1) = c(i)
5         Next
6         CollectionToColumn = Res

7         Exit Function
ErrHandler:
8         Throw "#CollectionToColumn (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LastElementOf
' Author     : Philip Swannell
' Date       : 18-Feb-2022
' Purpose    : Returns the last element of a collection
' -----------------------------------------------------------------------------------------------------------------------
Private Function LastElementOf(c As Collection)
1         On Error GoTo ErrHandler
2         LastElementOf = c(c.Count)
3         Exit Function
ErrHandler:
4         Throw "#LastElementOf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConstructExtraFxForwards
' Author     : Philip Swannell
' Date       : 18-Feb-2022
' Purpose    :
' Parameters :
'  ModelBareBones   :
'  ReceiveCurrency  :
'  PayCurrency      :
'  ReceiveNotionals :
'  MaturityDates    :
'  OnlyNonZeroTrades:
' -----------------------------------------------------------------------------------------------------------------------
Function ConstructExtraFxForwards(ModelBareBones As Dictionary, ReceiveCurrency As String, PayCurrency As String, _
          ReceiveNotionals, MaturityDates() As Date, OnlyNonZeroTrades As Boolean)

          Dim AnchorDate As Date
          Dim ForwardsDict As Dictionary
          Dim i As Long
          Dim k As Long
          Dim NTrades As Long
          Dim Trades() As Variant

          Const cnTradeID As Long = 1
          Const cnValuationFunction As Long = 2
          Const cnCounterparty As Long = 3
          Const cnStartDate As Long = 4
          Const cnEndDate As Long = 5
          Const cnReceiveCurrency As Long = 6
          Const cnPayCurrency As Long = 7
          Const cnReceiveNotional As Long = 9
          Const cnPayNotional As Long = 8

1         On Error GoTo ErrHandler

2         If OnlyNonZeroTrades Then
3             For i = 1 To sNRows(ReceiveNotionals)
4                 If ReceiveNotionals(i, 1) <> 0 Then
5                     NTrades = NTrades + 1
6                 End If
7             Next i
8         Else
9             NTrades = sNRows(ReceiveNotionals)
10        End If
11        AnchorDate = DictGet(ModelBareBones, "AnchorDate")
          
12        If Not ModelBareBones.Exists("EURUSDForwards") Then
13            Throw "Cannot find ""EURUSDForwards"" in ModelBareBones"
14        Else
15            Set ForwardsDict = DictGet(ModelBareBones, "EURUSDForwards")
16        End If
          
17        ReDim Trades(1 To NTrades + 1, 1 To 9)
18        Trades(1, cnTradeID) = "TradeID"
19        Trades(1, cnValuationFunction) = "ValuationFunction"
20        Trades(1, cnCounterparty) = "Counterparty"
21        Trades(1, cnStartDate) = "StartDate"
22        Trades(1, cnEndDate) = "EndDate"
23        Trades(1, cnReceiveCurrency) = "ReceiveCurrency"
24        Trades(1, cnPayCurrency) = "PayCurrency"
25        Trades(1, cnReceiveNotional) = "ReceiveNotional"
26        Trades(1, cnPayNotional) = "PayNotional"

          'Construct an on-market trade
27        k = 1
28        For i = 1 To sNRows(ReceiveNotionals)
29            If ReceiveNotionals(i, 1) <> 0 Or Not OnlyNonZeroTrades Then
30                k = k + 1
31                Trades(k, cnTradeID) = "ExtraTrade" & CStr(k - 1)
32                Trades(k, cnValuationFunction) = "FxForward"
33                Trades(k, cnCounterparty) = "Not Specified"
34                Trades(k, cnStartDate) = AnchorDate
35                Trades(k, cnEndDate) = MaturityDates(i, 1)

36                Trades(k, cnReceiveCurrency) = ReceiveCurrency
37                Trades(k, cnPayCurrency) = PayCurrency
38                Trades(k, cnReceiveNotional) = ReceiveNotionals(i, 1)
39                If Not ForwardsDict.Exists(MaturityDates(i, 1)) Then
40                    Throw "EURUSDForwards dictionary does not contain forward rate for " _
                          + Format(MaturityDates(i, 1), "yyyy-mmm-dd")
41                End If
42                Trades(k, cnPayNotional) = -ReceiveNotionals(i, 1) / ForwardsDict(MaturityDates(i, 1))
43            End If
44        Next i

45        ConstructExtraFxForwards = Trades

46        Exit Function
ErrHandler:
47        Throw "#ConstructExtraFxForwards (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ConstructExtraFxOptions(ModelBareBones As Dictionary, Ccy As String, _
          USDAmounts, MaturityDates() As Date, isCall As Boolean, OnlyNonZeroTrades As Boolean)
          
          'sign of USDAmounts indicates long or short the option from the Bank's pov
          'isCall = True if it's a call on CCy/Numeraire, put on Numeraire/CCy
          
          Const cnTradeID As Long = 1
          Const cnValuationFunction As Long = 2
          Const cnCounterparty As Long = 3
          Const cnEndDate As Long = 4
          Const cnCurrency As Long = 5
          Const cnNotional As Long = 6
          Const cnStrike As Long = 7
          Const cnIsCall As Long = 8
          
          Dim AnchorDate As Date
          Dim ForwardsDict As Dictionary
          Dim i As Long
          Dim k As Long
          Dim NTrades As Long
          Dim Trades() As Variant

1         On Error GoTo ErrHandler

2         If Ccy <> "USD" Then Throw "CCy must be USD"

3         If OnlyNonZeroTrades Then
4             For i = 1 To sNRows(USDAmounts)
5                 If USDAmounts(i, 1) <> 0 Then
6                     NTrades = NTrades + 1
7                 End If
8             Next i
9         Else
10            NTrades = sNRows(USDAmounts)
11        End If
12        AnchorDate = DictGet(ModelBareBones, "AnchorDate")
          
13        If Not ModelBareBones.Exists("EURUSDForwards") Then
14            Throw "Cannot find ""EURUSDForwards"" in ModelBareBones"
15        Else
16            Set ForwardsDict = DictGet(ModelBareBones, "EURUSDForwards")
17        End If
          
18        ReDim Trades(1 To NTrades + 1, 1 To 8)
19        Trades(1, cnTradeID) = "TradeID"
20        Trades(1, cnValuationFunction) = "ValuationFunction"
21        Trades(1, cnCounterparty) = "Counterparty"
22        Trades(1, cnEndDate) = "EndDate"
23        Trades(1, cnCurrency) = "Currency"
24        Trades(1, cnNotional) = "Notional"
25        Trades(1, cnStrike) = "Strike"
26        Trades(1, cnIsCall) = "IsCall"

          'Construct an atm trade
27        k = 1
28        For i = 1 To sNRows(USDAmounts)
29            If USDAmounts(i, 1) <> 0 Or Not OnlyNonZeroTrades Then
30                k = k + 1
31                Trades(k, cnTradeID) = "ExtraTrade" & CStr(k - 1)
32                Trades(k, cnValuationFunction) = "FxOption"
33                Trades(k, cnCounterparty) = "Not Specified"
34                Trades(k, cnEndDate) = MaturityDates(i, 1)

35                Trades(k, cnCurrency) = Ccy
36                Trades(k, cnNotional) = USDAmounts(i, 1)

37                If Not ForwardsDict.Exists(MaturityDates(i, 1)) Then
38                    Throw "EURUSDForwards dictionary does not contain forward rate for " _
                          + Format(MaturityDates(i, 1), "yyyy-mmm-dd")
39                End If
40                Trades(k, cnStrike) = 1 / ForwardsDict(MaturityDates(i, 1))
41                Trades(k, cnIsCall) = isCall
42            End If
43        Next i

44        ConstructExtraFxOptions = Trades

45        Exit Function
ErrHandler:
46        Throw "#ConstructExtraFxOptions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


