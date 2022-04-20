Attribute VB_Name = "modFunctions"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsCurrencySheet
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Returns TRUE if a worksheet (in the market data workbook) looks to be a
'             sheet containing rates and vol data for one currency
' -----------------------------------------------------------------------------------------------------------------------
Function IsCurrencySheet(ws As Worksheet) As Boolean
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = ws.Range("SwapRatesInit")
3         If Not R.Parent Is ws Then GoTo ErrHandler
4         Set R = ws.Range("XccyBasisSpreadsInit")
5         If Not R.Parent Is ws Then GoTo ErrHandler
6         Set R = ws.Range("VolInit")
7         If Not R.Parent Is ws Then GoTo ErrHandler
8         IsCurrencySheet = True
9         Exit Function
ErrHandler:
10        IsCurrencySheet = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsInflationSheet
' Author    : Philip Swannell
' Date      : 24-Apr-2017
' Purpose   : Returns TRUE if a worksheet (in the market data workbook) looks to be a
'             sheet containing data about an inflation index
' -----------------------------------------------------------------------------------------------------------------------
Function IsInflationSheet(ws As Worksheet) As Boolean
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = ws.Range("ZCSwapsInit")
3         If Not R.Parent Is ws Then GoTo ErrHandler
4         Set R = ws.Range("SeasonalAdjustments")
5         If Not R.Parent Is ws Then GoTo ErrHandler
6         Set R = ws.Range("HistoricDataInit")
7         If Not R.Parent Is ws Then GoTo ErrHandler
8         IsInflationSheet = True
9         Exit Function
ErrHandler:
10        IsInflationSheet = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sParseFrequencyString
' Author    : Philip Swannell
' Date      : 27-Apr-2016
' Purpose   : Convert a string description of payment frequency to a number
'             CHANGING THIS FUNCTION? Then also make equivalent change to method
'             SetValidationForFreq in SCRiPT.xlsm
' -----------------------------------------------------------------------------------------------------------------------
Function sParseFrequencyString(FrequencyString As String, ThrowOnError As Boolean, Optional ReturnNumber As Boolean = True) As Variant
1         On Error GoTo ErrHandler
          Const ErrString = "Frequency not recognised. Allowed values: Annual, Semi annual, Quarterly, Monthly. Can be abbreviated to first letter only."

2         Select Case LCase(FrequencyString)
          Case "annual", "ann", "a"
3             sParseFrequencyString = IIf(ReturnNumber, 1, "Annual")
4         Case "semi annual", "semi", "semi-annual", "s"
5             sParseFrequencyString = IIf(ReturnNumber, 2, "Semi annual")
6         Case "quarterly", "quarter", "q", "quart"
7             sParseFrequencyString = IIf(ReturnNumber, 4, "Quarterly")
8         Case "monthly", "month", "m"
9             sParseFrequencyString = IIf(ReturnNumber, 12, "Monthly")
10        Case Else
11            If ThrowOnError Then
12                Throw ErrString
13            Else
14                sParseFrequencyString = "#" + ErrString + "!"
15            End If
16        End Select
17        Exit Function
ErrHandler:
18        Throw "#sParseFrequencyString (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sParseDCT
' Author    : Philip Swannell
' Date      : 10-May-2016
' Purpose   : Validate and standardise a day count type
'             CHANGING THIS FUNCTION? Then also make equivalent change to method sSupportedDCTs
' -----------------------------------------------------------------------------------------------------------------------
Function sParseDCT(DCT As String, IsFloating As Boolean, ThrowOnError As Boolean)
          Dim AllowedForFloating As Boolean
          Dim ErrString As String
1         On Error GoTo ErrHandler

2         Select Case UCase(DCT)
          Case "A/360", "ACT/360", "ACTUAL/360"
3             AllowedForFloating = True
4             sParseDCT = "A/360"
5         Case "A/365F", "ACT/365F", "ACTUAL/365F", "ACT/365"
6             AllowedForFloating = True
7             sParseDCT = "A/365F"
8         Case "A/365L", "ACT/365L", "ACTUAL/365L", "ACT/ACT29"
9             sParseDCT = "Act/365L"
10        Case "30/360"
11            AllowedForFloating = True    'PGS 12 Jan 2017. Airbus have a USD trade with floating leg on 30/360 basis!!!
12            sParseDCT = UCase(DCT)
13        Case "30E/360", "30E/360 (ISDA)"
14            sParseDCT = UCase(DCT)
15        Case "ACT/ACT", "ACTUAL/ACTUAL"
16            sParseDCT = "Act/Act"
17        Case "ACTB/ACTB"
18            AllowedForFloating = True 'PGS 31 Jan 2019 Deloitte's client have floating legs on this basis
19            sParseDCT = "ActB/ActB"
20        Case Else
21            ErrString = "Invalid day count type: '" + DCT + "' Valid types are " + sConcatenateStrings(sSupportedDCTs())
22            If ThrowOnError Then
23                Throw ErrString
24            Else
25                sParseDCT = "#" + ErrString + "!"
26            End If
27        End Select

28        If IsFloating Then
29            If Not AllowedForFloating Then
30                ErrString = "Invalid day count type for floating leg: " + DCT + " Valid types are A/360, A/365F, 30/360, ActB/ActB"
31                If ThrowOnError Then
32                    Throw ErrString
33                Else
34                    sParseDCT = "#" + ErrString + "!"
35                End If
36            End If
37        End If

38        Exit Function
ErrHandler:
39        Throw "#sParseDCT (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSupportedDCTs
' Author    : Philip Swannell
' Date      : 09-May-2016
' Purpose   : Returns a list of the supported day count types. This function needs
'             to be kept in synch with the R function DCF
'       CHANGING THIS FUNCTION? Then also make equivalent change to method sParseDCT
' -----------------------------------------------------------------------------------------------------------------------
Function sSupportedDCTs()
          Dim Res() As String
1         ReDim Res(1 To 8, 1 To 1)
2         Res(1, 1) = "30/360"
3         Res(2, 1) = "30E/360"
4         Res(3, 1) = "30E/360 (ISDA)"
5         Res(4, 1) = "A/360"
6         Res(5, 1) = "A/365F"
7         Res(6, 1) = "Act/Act"
8         Res(7, 1) = "Act/365L"
9         Res(8, 1) = "ActB/ActB"
10        sSupportedDCTs = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCurrencies
' Author    : Philip Swannell
' Date      : 10-Oct-2016
' Purpose   : Returns lists of currencies, three letter ISO codes or ISO codes with annotation
' -----------------------------------------------------------------------------------------------------------------------
Function sCurrencies(Optional LongForm As Boolean = False, Optional MainCurrenciesOnly As Boolean = True)
1         On Error GoTo ErrHandler
          Dim AllCurrencies
          Dim AllLongForms
          Dim ChooseVector

2         With RangeFromSheet(shSAIStaticData, "AllCurrencies")
3             ChooseVector = .Columns(1).Value
4             AllCurrencies = .Columns(2).Value
5             AllLongForms = .Columns(3).Value
6         End With

7         If MainCurrenciesOnly Then
8             If LongForm Then
9                 sCurrencies = sArrayConcatenate(sMChoose(AllCurrencies, ChooseVector), " - ", _
                                                  sMChoose(AllLongForms, ChooseVector))
10            Else
11                sCurrencies = sMChoose(AllCurrencies, ChooseVector)
12            End If
13        Else
14            If LongForm Then
15                sCurrencies = sArrayConcatenate(AllCurrencies, " - ", AllLongForms)
16            Else
17                sCurrencies = AllCurrencies
18            End If
19        End If

20        Exit Function
ErrHandler:
21        Throw "#sCurrencies (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SupportedInflationIndices
' Author    : Philip Swannell
' Date      : 21-Apr-2017
' Purpose   : Get a list of the valid inflation indices - for pick list and trade validation...
' -----------------------------------------------------------------------------------------------------------------------
Function SupportedInflationIndices()
1         On Error GoTo ErrHandler
2         SupportedInflationIndices = RangeFromSheet(shSAIStaticData, "InflationIndices").Columns(1).Value
3         Exit Function
ErrHandler:
4         Throw "#SupportedInflationIndices (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InflationIndexInfo
' Author    : Philip Swannell
' Date      : 12-May-2017
' Purpose   : Code up these properties only here...
' -----------------------------------------------------------------------------------------------------------------------
Function InflationIndexInfo(Index As String, Info As String)
1         On Error GoTo ErrHandler
2         Select Case LCase(Info)
          Case "basecurrency", "base currency"
3             InflationIndexInfo = sVlookup(Index, shSAIStaticData.Range("InflationIndices").Value, 4)
4         Case "lag"
5             InflationIndexInfo = sVlookup(Index, shSAIStaticData.Range("InflationIndices").Value, 3)
6         Case "description"
7             InflationIndexInfo = sVlookup(Index, shSAIStaticData.Range("InflationIndices").Value, 2)
8         Case Else
9             Throw "Info not recognised. Must be one of: 'BaseCurrency', 'Lag', 'Description'"
10        End Select
11        Exit Function
ErrHandler:
12        InflationIndexInfo = "#InflationIndexInfo (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FirstElementOf
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : sFns.Providers.Default.EvaluateExpression seems inconsistent in the number
'             of dimensions of the data returned. This function returns the first element
'             of the return for 0 to 3 dimensions.
' -----------------------------------------------------------------------------------------------------------------------
Function FirstElementOf(This)
1         On Error GoTo ErrHandler
2         Select Case NumDimensions(This)

          Case 0
3             FirstElementOf = This
4         Case 1
5             FirstElementOf = This(LBound(This))
6         Case 2
7             FirstElementOf = This(LBound(This, 1), LBound(This, 2))
8         Case 3
9             FirstElementOf = This(LBound(This, 1), LBound(This, 2), LBound(This, 3))
10        Case Else
11            Throw "Unexpected error - array has more than three dimensions"
12        End Select
13        Exit Function
ErrHandler:
14        Throw "#FirstElementOf (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CurrenciesFromQuery
' Author    : Philip Swannell
' Date      : 28-Jul-2016
' Purpose   : Given a query onto the trades held in the trades workbook, what currencies
'             are involved? Return is column array, sorted alphabetically.
' -----------------------------------------------------------------------------------------------------------------------
Function CurrenciesFromQuery(FilterBy1 As String, Filter1Value, FilterBy2 As String, Filter2Value, _
          IncludeFutureTrades As Boolean, PortfolioAgeing As Double, WithFxTrades As Boolean, _
          WithRatesTrades As Boolean, twb As Workbook, _
          FutureTradesSheet As Worksheet, AnchorDate As Date)

          Dim RatesCcys2
          Dim RatesCcys1
          Dim FxCcys1
          Dim FxCcys2
          Dim ChooseVector
          Dim FilteringNecessary As Boolean
          Dim TC As TradeCount
          Dim UseCSV As Boolean
          Dim HeaderCcy1 As String
          Dim HeaderCcy2 As String
          
1         On Error GoTo ErrHandler

2         UseCSV = IsTradesWorkbook2022Style(twb)
          
3         FilteringNecessary = Not (LCase(FilterBy1) = "none" And LCase(FilterBy2) = "none")

4         If WithFxTrades Then
5             FxCcys1 = GetColumnFromTradesWorkbook(IIf(UseCSV, "Prim Cur", "CCY1"), IncludeFutureTrades, PortfolioAgeing, True, False, twb, FutureTradesSheet, AnchorDate)
6             FxCcys2 = GetColumnFromTradesWorkbook(IIf(UseCSV, "Sec Cur", "CCY2"), IncludeFutureTrades, PortfolioAgeing, True, False, twb, FutureTradesSheet, AnchorDate)
7             If FilteringNecessary Then
8                 ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, PortfolioAgeing, False, True, False, "All", TC, twb, FutureTradesSheet, AnchorDate)
9                 If sColumnOr(ChooseVector)(1, 1) Then
10                    FxCcys1 = sMChoose(FxCcys1, ChooseVector)
11                    FxCcys2 = sMChoose(FxCcys2, ChooseVector)
12                Else
13                    FxCcys1 = CreateMissing()
14                    FxCcys2 = CreateMissing()
15                End If
16            End If
17        Else
18            FxCcys1 = CreateMissing()
19            FxCcys2 = CreateMissing()
20        End If

21        If WithRatesTrades Then
22            RatesCcys1 = GetColumnFromTradesWorkbook(IIf(UseCSV, "Rec Ccy", "CCY_REC"), IncludeFutureTrades, PortfolioAgeing, False, True, twb, FutureTradesSheet, AnchorDate)
23            RatesCcys2 = GetColumnFromTradesWorkbook(IIf(UseCSV, "Pay Ccy", "CCY_PAY"), IncludeFutureTrades, PortfolioAgeing, False, True, twb, FutureTradesSheet, AnchorDate)
24            If FilteringNecessary Then
25                ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, PortfolioAgeing, False, False, True, "All", TC, twb, FutureTradesSheet, AnchorDate)
26                If sColumnOr(ChooseVector)(1, 1) Then
27                    RatesCcys1 = sMChoose(RatesCcys1, ChooseVector)
28                    RatesCcys2 = sMChoose(RatesCcys2, ChooseVector)
29                Else
30                    RatesCcys1 = CreateMissing()
31                    RatesCcys2 = CreateMissing()
32                End If
33            End If
34        Else
35            RatesCcys1 = CreateMissing()
36            RatesCcys2 = CreateMissing()
37        End If

38        CurrenciesFromQuery = sSortedArray(sRemoveDuplicates(sArrayStack(FxCcys1, FxCcys2, RatesCcys1, RatesCcys2), True))

39        Exit Function
ErrHandler:
40        CurrenciesFromQuery = "#CurrenciesFromQuery (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddCayleyFiltersToMRU
' Author    : Philip Swannell
' Date      : 21-Sep-2016
' Purpose   : In Cayley workbook, works in conjuction with the PFE sheet double-click event to present
'             "Most Recently Used" values for the Filter1Value and Filter2Value fields
'             Also called from method ShowSelectTrades in this project.
' -----------------------------------------------------------------------------------------------------------------------
Sub AddCayleyFiltersToMRU(twb As Workbook, FilterBy1 As String, Filter1Value As Variant, FilterBy2 As String, _
    Filter2Value As Variant, AnchorDate As Date)
    
          Dim ChooseVector As Variant
          Dim Filter1ValueIsValid As Boolean
          Dim Filter2ValueIsValid As Boolean
          Dim FilterBy1IsValid As Boolean
          Dim FilterBy2IsValid As Boolean
          Dim ModelName As String
          Dim TC As TradeCount
          Dim TradeHeaders As Variant

1         On Error GoTo ErrHandler

2         ModelName = "CayleyModel"

3         TradeHeaders = GetColumnFromTradesWorkbook("AllHeaders", False, 0, True, True, twb, twb.Worksheets(1), AnchorDate)
4         FilterBy1IsValid = IsNumeric(sMatch(FilterBy1, TradeHeaders))
5         FilterBy2IsValid = IsNumeric(sMatch(FilterBy2, TradeHeaders))

6         If Not FilterBy1IsValid Then
7             Filter1ValueIsValid = False
8         Else
9             Select Case VarType(Filter1Value)
                  Case vbBoolean, vbDouble, vbInteger, vbLong
10                    Filter1ValueIsValid = True
11                Case vbString
12                    Filter1ValueIsValid = VarType(sIsRegMatch(CStr(Filter1Value), "Foo")) = vbBoolean
13                Case Else
14                    Filter1ValueIsValid = False
15            End Select
16            If Filter1ValueIsValid Then
17                If Filter1Value <> "None" Then
                      'Check that the filter selects some but not all of the trades
18                    ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, "None", "None", False, 0, False, True, True, "All", TC, twb, twb.Worksheets(1), AnchorDate)
19                    Filter1ValueIsValid = IsNumber(sMatch(True, ChooseVector)) And (IsNumber(sMatch(False, ChooseVector)))
20                End If
21            End If
22        End If

23        If Not FilterBy2IsValid Then
24            Filter2ValueIsValid = False
25        Else
26            Select Case VarType(Filter2Value)
                  Case vbBoolean, vbDouble, vbInteger, vbLong
27                    Filter2ValueIsValid = True
28                Case vbString
29                    Filter2ValueIsValid = VarType(sIsRegMatch(CStr(Filter2Value), "Foo")) = vbBoolean
30                Case Else
31                    Filter2ValueIsValid = False
32            End Select
33            If Filter2ValueIsValid Then
34                If Filter2Value <> "None" Then
                      'Check that the filter selects some but not all of the trades
35                    ChooseVector = ChooseVectorFromFilters("None", "None", FilterBy2, Filter2Value, False, 0, False, True, True, "All", TC, twb, twb.Worksheets(1), AnchorDate)
36                    Filter2ValueIsValid = IsNumber(sMatch(True, ChooseVector)) And (IsNumber(sMatch(False, ChooseVector)))
37                End If
38            End If
39        End If

40        If Filter1ValueIsValid Then AddFilterToMRU "CayleyFilterBy" & FilterBy1, CStr(Filter1Value)
41        If Filter2ValueIsValid Then AddFilterToMRU "CayleyFilterBy" & FilterBy2, CStr(Filter2Value)
42        If FilterBy1IsValid Then AddFilterToMRU "CayleyFilterBy1", CStr(FilterBy1)
43        If FilterBy2IsValid Then AddFilterToMRU "CayleyFilterBy2", CStr(FilterBy2)

44        Exit Sub
ErrHandler:
45        Throw "#AddCayleyFiltersToMRU (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsTradesWorkbook2022Style
' Author     : Philip Swannell
' Date       : 04-Mar-2022
' Purpose    : Test if a trades workbook is 2022-style (i.e. an in-memory cache of the data provided in .csv files) or
'              2017 style, a workbok that Airbus was constructed themselves.
' -----------------------------------------------------------------------------------------------------------------------
Function IsTradesWorkbook2022Style(twb As Workbook) As Boolean
1         On Error GoTo ErrHandler
2         If IsInCollection(twb.Worksheets, SN_FxTrades2) Then
3             IsTradesWorkbook2022Style = True
4         ElseIf IsInCollection(twb.Worksheets, SN_FxTrades) Then
5             IsTradesWorkbook2022Style = False
6         Else
7             Throw "Unexpected error: Cannot find worksheet for Fx trades in workbook '" + twb.Name + "' " & _
              "we looked for worksheets named '" & SN_FxTrades & "' or '" & SN_FxTrades2 & "'"
8         End If
9         Exit Function
ErrHandler:
10        Throw "#IsTradesWorkbook2022Style (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckTradesWorkbook
' Author    : Philip Swannell
' Date      : 03-Dec-2016
' Purpose   : Check the trades workbook when we open it to throw helpful error message if
'             something is wrong with it. Argument CalledFromCayley is used only for
'             generation of any error message.
' -----------------------------------------------------------------------------------------------------------------------
Function CheckTradesWorkbook(wb As Workbook, CalledFromCayley As Boolean, ThrowErrors As Boolean)

          Dim Postamble As String
          Dim Preamble As String
          Const FxHeadersNeeded = "CCY1,CCY2,CPTY_PARENT,DEAL_TYPE,FWD_1,FWD_2,MATURITY_DATE,OP_FINANCE,VALUE_DATE"
          Const RatesHeadersNeeded = "BASIS_PAY,BASIS_REC,CCY_PAY,CCY_REC,CPTY_PARENT,DEAL_TYPE,INDEX_PAY,INDEX_REC,MATURITY_DATE,NOMINAL_PAY,NOMINAL_REC,OP_FINANCE,PAY_COUPON_DTROLL,PAY_COUPON_FREQ,REC_COUPON_DTROLL,REC_COUPON_FREQ,SPREAD_PAY,SPREAD_REC,VALUE_DATE"
          Const AmortHeadersNeeded = "TRADE_ID,START_DATE,PAY_REC_LEG,NOTIONAL"
          Dim CompareRes
          Dim CopyOfErr As String
          Dim HaveHeaders
          Dim i As Long
          Dim NeedHeaders
          Dim SheetName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         If CalledFromCayley Then
3             Preamble = "There is a problem with the trades workbook '" + wb.Name + "':" + vbLf + vbLf
4             Postamble = vbLf + vbLf + "You cannot use this trades workbook." + vbLf + vbLf + "Possible causes:" + vbLf + _
                          "a) The 'Trades Workbook' setting on the Config sheet is not correct or points to the wrong workbook." + vbLf + _
                          "b) You are trying to use an out-of-date trades workbook which does not contain all the data that's now needed." + vbLf + _
                          "c) There was an error in the system that generated the trades workbook."
5         Else
6             Preamble = "There is a problem with the trades file '" + wb.Name + "':" + vbLf + vbLf
7             Postamble = vbLf + vbLf + "You cannot use this file." + vbLf + vbLf + "Possible causes:" + vbLf + _
                          "a) You have opened a file that's not a trades file." + vbLf + _
                          "b) You are trying to use an out-of-date file which does not contain all the data that's now needed." + vbLf + _
                          "c) There was an error in the system that generated the file."
8         End If

9         If Not IsInCollection(wb.Worksheets, SN_FxTrades) Then
10            Throw Preamble + "There is no worksheet '" + SN_FxTrades + "'" + Postamble
11        ElseIf Not IsInCollection(wb.Worksheets, SN_RatesTrades) Then
12            Throw Preamble + "There is no worksheet '" + SN_RatesTrades + "'" + Postamble
13        ElseIf Not IsInCollection(wb.Worksheets, SN_Amortisation) Then
14            Throw Preamble + "There is no worksheet '" + SN_Amortisation + "'" + Postamble
15        End If

16        For i = 1 To 3
17            SheetName = Choose(i, SN_FxTrades, SN_RatesTrades, SN_Amortisation)
18            Set ws = wb.Worksheets(SheetName)
19            Select Case ws.ListObjects.Count
              Case 1
20                HaveHeaders = sArrayTranspose(ws.ListObjects(1).HeaderRowRange.Value)
21            Case 0
                  'Not in a table, so assume data starts at A1
22                If IsEmpty(ws.Range("A1")) Then Throw " cannot find data on worksheet '" + SheetName + "'" + Postamble
23                HaveHeaders = sArrayTranspose(sExpandRight(ws.Range("A1")))
24            Case Is > 1
25                Throw Preamble + " The worksheet '" + SheetName + "' contains more than one Table" + Postamble
26            End Select

27            NeedHeaders = sTokeniseString(Choose(i, FxHeadersNeeded, RatesHeadersNeeded, AmortHeadersNeeded))
28            CompareRes = sCompareTwoArrays(NeedHeaders, HaveHeaders, "In1AndNotIn2")
29            If sNRows(CompareRes) > 1 Then
30                CompareRes = sConcatenateStrings(sSubArray(CompareRes, 2, 1, , 1), ", ")
31                Throw Preamble + "The data on the worksheet '" + SheetName + "' does not have the following headers: " + CompareRes + Postamble
32            End If
33        Next i

34        CheckTradesWorkbook = "OK"

35        Exit Function
ErrHandler:
36        CopyOfErr = "#CheckTradesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
37        If ThrowErrors Then
38            Throw CopyOfErr, True
39        Else
40            CheckTradesWorkbook = CopyOfErr
41        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckMarketWorkbook
' Author    : Philip Swannell
' Date      : 03-Dec-2016
' Purpose   : Produce a friendly error message if the market data workbook look wrong or out-of-date
' -----------------------------------------------------------------------------------------------------------------------
Function CheckMarketWorkbook(wb As Workbook, NameOfCallingBook As String)

          Const MINVERSIONMDW = 172       'correct as of 27 Nov 2017
          Dim HaveVersionCaller As Long
          Dim HaveVersionMDW As Long
          Dim Postamble As String
          Dim Preamble As String
          Const SheetsNeeded = "Audit,FX,Credit,HistoricalCorrEUR,Config,EUR,USD,GBP"
          Dim i As Long
          Dim SheetList
1         On Error GoTo ErrHandler

2         SheetList = sTokeniseString(SheetsNeeded)

3         Preamble = "There is a problem with the market data workbook '" + wb.Name + "':" + vbLf + vbLf
4         Postamble = vbLf + vbLf + "You cannot use this market data workbook." + vbLf + vbLf + "Possible causes:" + vbLf + _
                      "a) The 'MarketDataWorkbook' setting on the Config sheet is not correct or points to the wrong workbook." + vbLf + _
                      "b) You are trying to use an out-of-date MarketDataWorkbook."

5         For i = 1 To sNRows(SheetList)
6             If Not IsInCollection(wb.Worksheets, CStr(SheetList(i, 1))) Then
7                 Throw Preamble + "There is no worksheet '" + CStr(SheetList(i, 1)) + "'" + Postamble
8             End If
9         Next i

10        HaveVersionMDW = RangeFromSheet(wb.Worksheets("Audit"), "Headers").Cells(2, 1).Value
11        HaveVersionCaller = RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value
12        If HaveVersionMDW < MINVERSIONMDW Then
13            Throw Preamble + "It is version " + CStr(HaveVersionMDW) + " but to be compatible with this version of the " + NameOfCallingBook + " workbook, you need version " + CStr(MINVERSIONMDW) + " or later. The version numbers for both workbooks are shown on their Audit sheets." + Postamble
14        End If

15        Exit Function
ErrHandler:
16        Throw "#CheckMarketWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!", True
End Function


Sub TestCheckLinesWorkbook()
1         On Error GoTo ErrHandler
2         CheckLinesWorkbook ActiveWorkbook, True, False

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestCheckLinesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub



' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckLinesWorkbook
' Author    : Philip Swannell
' Date      : 03-Dec-2016
' Purpose   : Check the lines workbook when we open it to throw helpful error message if
'             something is wrong with it
' -----------------------------------------------------------------------------------------------------------------------
Function CheckLinesWorkbook(wb As Workbook, ForCredit As Boolean, ForCapital As Boolean)

          Dim Postamble As String
          Dim Preamble As String
          Const CreditHeadersNeeded = "Very short name,CPTY LONG NAME,CPTY_PARENT,METHODOLOGY,Confidence %,Shortfall or Quantile,Volatility Input,Base Currency,Notional Cap,Product Credit Limits,Fx Notional Weights,Rates Notional Weights,Line Interp.,1Y Limit,2Y Limit,3Y Limit,4Y Limit,5Y Limit,7Y Limit,10Y Limit,THR Bank 1Y,THR Bank 3Y,Airbus THR 3Y"
          Const CapitalHeadersNeeded = "CPTY_PARENT,Risk Weight Method,EaD/CVA Capital Charge Methods,Exemption,Capital Hurdle,Capital Discount,Spread Vol,Stressed Spread Vol,Portfolio Size,Index Correlation,PD,Recovery,Alpha,Hedge Offset,DVA benefit %,FVA charge %"
          Dim CompareRes
          Dim HaveHeaders
          Dim NeedHeaders
          Dim SheetName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Preamble = "There is a problem with the lines workbook '" + wb.Name + "':" + vbLf + vbLf
3         Postamble = vbLf + vbLf + "You cannot use this lines workbook." + vbLf + vbLf + "Possible causes:" + vbLf + _
              "a) The 'Lines Workbook' setting on the Config sheet is not correct or points to the wrong workbook." + vbLf + _
              "b) You are trying to use an out-of-date lines workbook which does not contain all the data that's needed." + vbLf + _
              "c) There was an error in the system that generated the lines workbook."

4         If Not IsInCollection(wb.Worksheets, SN_Lines) Then
5             Throw Preamble + "There is no worksheet '" + SN_Lines + "'" + Postamble
6         End If

7         If ForCredit And ForCapital Then
8             NeedHeaders = sRemoveDuplicateRows(sArrayStack(sTokeniseString(CreditHeadersNeeded), sTokeniseString(CapitalHeadersNeeded)))
9         ElseIf ForCredit Then
10            NeedHeaders = sTokeniseString(CreditHeadersNeeded)
11        ElseIf ForCapital Then
12            NeedHeaders = sTokeniseString(CapitalHeadersNeeded)
13        Else
14            Throw "At least one of ForCapital and ForCredit must be true"
15        End If

16        SheetName = SN_Lines
17        Set ws = wb.Worksheets(SheetName)
18        Select Case ws.ListObjects.Count
              Case 0
19                Throw Preamble + " The worksheet '" + SheetName + "' does not contain a Table (in the sense of Excel Ribbon > Insert > Table)" & Postamble
20            Case Is > 1
21                Throw Preamble & " The worksheet '" & SheetName & "' contains more than one Table" & Postamble
22        End Select
23        HaveHeaders = sArrayTranspose(ws.ListObjects(1).HeaderRowRange.Value)
24        CompareRes = sCompareTwoArrays(NeedHeaders, HaveHeaders, "In1AndNotIn2")
25        If sNRows(CompareRes) > 1 Then
26            CompareRes = sConcatenateStrings(sSubArray(CompareRes, 2, 1, , 1), ", ")
27            Throw Preamble & "The table '" & ws.ListObjects(1).Name & "' on the worksheet '" & SheetName & "' does not have the following headers: " & CompareRes & Postamble
28        End If

          Dim RepeatedBanks As String, ColNum, AllBanks
29        ColNum = sMatch("CPTY_PARENT", HaveHeaders)

30        AllBanks = ws.ListObjects(1).DataBodyRange.Columns(ColNum).Value

31        RepeatedBanks = RepeatsInVector(AllBanks)

32        If Len(RepeatedBanks) > 0 Then
33            Throw Preamble & "The table '" & ws.ListObjects(1).Name & "' on the worksheet '" & SheetName & "' lists the following banks more than once: " & RepeatedBanks & vbLf & "Please correct this before proceeding." & Postamble
34        End If

35        Exit Function
ErrHandler:
36        Throw "#CheckLinesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!", True
End Function

Private Function RepeatsInVector(Elements) As String

          Dim ElementsSorted, ChooseVector, anyRepeats As Boolean, i As Long

1         On Error GoTo ErrHandler
2         ElementsSorted = sSortedArray(Elements)
3         ChooseVector = sReshape(False, sNRows(Elements), 1)
4         For i = 2 To sNRows(ElementsSorted)
5             If LCase(ElementsSorted(i, 1)) = LCase(ElementsSorted(i - 1, 1)) Then
6                 anyRepeats = True
7                 ChooseVector(i, 1) = True
8             End If
9         Next i
10        If Not anyRepeats Then
11            RepeatsInVector = ""
12        Else
13            RepeatsInVector = sConcatenateStrings(sMChoose(ElementsSorted, ChooseVector))
14        End If

15        Exit Function
ErrHandler:
16        Throw "#RepeatsInVector (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



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

1         On Error GoTo ErrHandler
2         AllNames = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
3         AllLongNames = GetColumnFromLinesBook("CPTY LONG NAME", LinesBook)

4         If ForCommandBar Then
5             AllPrettyNames = sJustifyArrayOfStrings(sArrayRange(AllNames, AllLongNames), "Segoe UI", 9, "           " & vbTab)
6         Else
7             AllPrettyNames = sJustifyArrayOfStrings(sArrayRange(AllNames, AllLongNames), "Tahoma", 8, " " & vbTab)
8         End If

9         If Annotate Then
10            Res = sVlookup(TheBanks, sArrayRange(AllNames, AllPrettyNames))
11        Else
12            Res = sVlookup(TheBanks, sArrayRange(AllPrettyNames, AllNames))
13        End If

14        Res = sArrayIf(sArrayEquals(Res, "#Not found!"), TheBanks, Res)

15        AnnotateBankNames = Res
16        Exit Function
ErrHandler:
17        Throw "#AnnotateBankNames (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetColumnFromLinesBook
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Returns the contents of an entire column from the Lines book
' -----------------------------------------------------------------------------------------------------------------------
Function GetColumnFromLinesBook(ByVal Header As String, LinesBook As Workbook)
1         On Error GoTo ErrHandler

          Dim ColNumber As Variant
          Dim EntireRange As Range
          Dim EntireRangeNoHeaders As Range
          Dim HeaderRow As Range
          Dim HeaderRowTranspose

2         GetLinesRanges EntireRange, EntireRangeNoHeaders, HeaderRow, LinesBook
3         HeaderRowTranspose = Application.WorksheetFunction.Transpose(HeaderRow.Value2)
4         ColNumber = sMatch(Header, HeaderRowTranspose)
5         If Not IsNumber(ColNumber) Then Throw "Cannot find column headed '" + Header + "' in the Lines Workbook"
6         GetColumnFromLinesBook = EntireRangeNoHeaders.Columns(ColNumber).Value

7         Exit Function
ErrHandler:
8         Throw "#GetColumnFromLinesBook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetLinesRanges
' Author    : Philip Swannell
' Date      : 09-Nov-2016
' Purpose   : Subroutine shared between LookupCounterpartyInfo and GetColumnFromLinesBook
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLinesRanges(ByRef EntireRange As Range, ByRef EntireRangeNoHeaders As Range, ByRef HeaderRange As Range, LinesBook As Workbook)
          Dim SheetName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         SheetName = SN_Lines

3         If Not IsInCollection(LinesBook.Worksheets, SheetName) Then
4             Throw "Lines workbook (" + LinesBook.Name + ") has no worksheet named " + SheetName
5         End If
6         Set ws = LinesBook.Worksheets(SheetName)
          Dim lo As ListObject

7         Select Case ws.ListObjects.Count
          Case 0
8             Throw SN_Lines + " in workbook " + LinesBook.Name + " must have a 'Table' (Ribbon > Insert > Table) containing the lines data."
9         Case 1
10            Set lo = ws.ListObjects(1)
11            Set EntireRangeNoHeaders = lo.DataBodyRange
12            Set HeaderRange = lo.HeaderRowRange
13            Set EntireRange = Application.Union(HeaderRange, EntireRangeNoHeaders)
14        Case Else
15            Throw SN_Lines + " in workbook " + LinesBook.Name + " must have just one 'Table' (Ribbon > Insert > Table) but it has " + CStr(ws.ListObjects.Count)
16        End Select
17        Exit Function
ErrHandler:
18        Throw "#GetLinesRanges (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSupportedInflationLagMethods
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : So we code up the allowed strings in only one place...
' -----------------------------------------------------------------------------------------------------------------------
Function sSupportedInflationLagMethods()
          Dim Res() As String
1         ReDim Res(1 To 3, 1 To 1)
2         Res(1, 1) = "2m no interpolation"
3         Res(2, 1) = "3m no interpolation"
4         Res(3, 1) = "Linear interpolation between 2m and 3m"
5         sSupportedInflationLagMethods = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSupportedInflationEffectiveDates
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : So we code up the allowed strings in only one place...
' -----------------------------------------------------------------------------------------------------------------------
Function sSupportedInflationEffectiveDates()
          Dim Res() As String
1         ReDim Res(1 To 2, 1 To 1)
2         Res(1, 1) = "T+2"
3         Res(2, 1) = "15th of Month"
4         sSupportedInflationEffectiveDates = Res
End Function


