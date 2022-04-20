Attribute VB_Name = "modSolvers"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FlushStatics
' Author    : Philip Swannell
' Date      : 05-Oct-2016
' Purpose   : Ensure that we clear out cached return in method GetTradesInJuliaFormat
' -----------------------------------------------------------------------------------------------------------------------
Sub FlushStatics()
          Dim Res
1         On Error GoTo ErrHandler
          Dim TC As TradeCount
          Dim twb As Workbook
          
2         On Error Resume Next
          'Call below will error, and the error handler resets static variables
3         Res = GetTradesInJuliaFormat("Foo", "Foo", "Foo", "Foo", False, 100, False, "EUR", False, False, 1, _
              "EUR,USD", False, TC, twb, shFutureTrades, Date)
              
4         Exit Sub
ErrHandler:
5         Throw "#FlushStatics (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PFEProfileHWFromFilters
' Author    : Philip Swannell
' Date      : 18-Jul-2016
' Purpose   : Get trades from the trades workbook, convert them to the format needed by
'             the Julia code, compress them and calculate the PFE profile.
' -----------------------------------------------------------------------------------------------------------------------
Function PFEProfileHWFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
          IncludeFutureTrades As Boolean, WithFxTrades As Boolean, WithRatesTrades As Boolean, _
          WithExtraTrades As Boolean, ExtraTradeLabels, ExtraTradeAmounts, PortfolioAgeing As Double, _
          FlipTrades As Boolean, Numeraire As String, Model As String, NumberOfSims As Long, _
          TimeGap As Double, TimeEnd As Double, PFEPercentile As Double, IsShortfall As Boolean, _
          ReportCurrency As String, TradesScaleFactor As Double, CurrenciesToInclude As String, _
          TC As TradeCount, twb As Workbook, CompressTrades As Boolean, ModelBareBones As Dictionary, _
          CalcByProduct As Boolean, ExtraTradesAre As String)
          
          Dim AnchorDate As Date
          Dim TheExtraTrades
          Dim TheTrades As Variant

1         On Error GoTo ErrHandler

2         AnchorDate = ModelBareBones("AnchorDate")

3         TheTrades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, FlipTrades, Numeraire, WithFxTrades, WithRatesTrades, _
              TradesScaleFactor, CurrenciesToInclude, CompressTrades, TC, twb, shFutureTrades, AnchorDate)

4         If WithExtraTrades Then
5             TheExtraTrades = ConstructExtraTrades(ExtraTradesAre, ModelBareBones, _
                  ExtraTradeAmounts, ExtraTradeLabels)
6         End If

7         PFEProfileHWFromFilters = PFEProfileHW(Model, NumberOfSims, TheTrades, TheExtraTrades, TimeGap, _
              TimeEnd, PFEPercentile, IsShortfall, ReportCurrency, CalcByProduct)
8         Exit Function
ErrHandler:
9         PFEProfileHWFromFilters = "#PFEProfileHWFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
13    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetBooleans
' Author    : Philip Swannell
' Date      : 21-Sep-2016
' Purpose   : Encapsulate deriving Booleans from two strings: IncludeAssetClasses and ProductCreditLimits
' -----------------------------------------------------------------------------------------------------------------------
Function SetBooleans(IncludeAssetClasses As String, ProductCreditLimits As String, ExtraTradesAre As String, _
          IncludeExtraTrades As Boolean, _
          ByRef IncludeFxTrades As Boolean, ByRef IncludeRatesTrades As Boolean, ByRef CalcByProduct)
1         On Error GoTo ErrHandler

2         If IncludeExtraTrades Then
              Dim ExtraTradesAreFx As Boolean
              Dim ExtraTradesAreRates As Boolean

3             ParseExtraTradesAre ExtraTradesAre, ExtraTradesAreFx, ExtraTradesAreRates

4         Else
5             ExtraTradesAreFx = False
6             ExtraTradesAreRates = False
7         End If

8         Select Case LCase(IncludeAssetClasses)
              Case "fx"
9                 IncludeFxTrades = True: IncludeRatesTrades = False
10            Case "rates"
11                IncludeFxTrades = False: IncludeRatesTrades = True
12            Case "rates and fx"
13                IncludeFxTrades = True: IncludeRatesTrades = True
14            Case Else
15                Throw "IncludeAssetClasses must be one of 'Rates', 'Fx' or 'Rates and Fx'"
16        End Select

17        Select Case ProductCreditLimits
              Case "Calc & Limit by Product", "Calculation & Limit by Product"
                  'Normal but we have to ensure that the limits entered are for Fx only
18                IncludeRatesTrades = False
19                CalcByProduct = False
20                If Not IncludeFxTrades Then
21                    Throw "For this bank, ProductCreditLimits entered in the lines workbook is '" + _
                          ProductCreditLimits + "' and therefore the credit limits (also entered in the " + _
                          "lines workbook) are assumed to be for the bank's Fx book only." + vbLf + vbLf + "So IncludeAssetClasses " + _
                          "cannot take the value '" + IncludeAssetClasses + "' because then there " & _
                          "will definitely be no Fx trades to count against the credit limits.", True
22                ElseIf ExtraTradesAreRates Then
23                    Throw "For this bank, ProductCreditLimits entered in the lines workbook is '" + _
                          ProductCreditLimits + "' and therefore the credit limits (also entered in the " + _
                          "lines workbook) are assumed to be for the bank's Fx book only." + vbLf + vbLf + "So ExtraTradesAre " + _
                          "cannot take the value '" + ExtraTradesAre + "' because those are Rates trades which the " + _
                          "bank trades under a separate limit", True
24                End If
                  
25            Case "Global Calc", "Global Calculation"
26                CalcByProduct = False
27            Case "Global Limit & Calc by Product", "Global Limit & Calculation by Product"
28                CalcByProduct = True
29            Case Else
30                Throw "ProductCreditLimits must be one of: " & _
                      " 'Global Calc', 'Calc & Limit by Product','Global Limit & Calc by Product'"
31        End Select

32        Exit Function
ErrHandler:
33        Throw "#SetBooleans (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

'TODO This method partly duplicates CheckCreditLimits. Rationalise
Sub ValidateCreditInputs(ShortfallOrQuantie As String, CreditLimits, CreditInterpMethod As String)

          Dim AnyPositive As Boolean
          Dim i As Long
          Dim NR As Long
1         On Error GoTo ErrHandler

2         Select Case CreditInterpMethod
              Case "Linear", "FlatToRight"
3             Case " - "
4                 Throw "For Headroom calculations to work 'FilterBy1' must be 'Counterparty Parent' and 'Filter1Value' must be the name of a Counterparty Parent from the Lines workbook (in column CPTY_PARENT)"
5             Case Else
6                 Throw "CreditInterpMathod must be 'Linear' or 'FlatToRight' but got '" & CStr(CreditInterpMethod) & "'"
7         End Select

8         NR = sNRows(CreditLimits)

9         Select Case ShortfallOrQuantie
              Case "Shortfall", "Quantile"
10            Case Else
11                Throw "ShortfallOrQuantile must be 'Shortfall' or 'Quantile' but got '" & CStr(ShortfallOrQuantie) & "'"
12        End Select

13        For i = 1 To NR
14            If Not IsNumber(CreditLimits(i, 1)) Then Throw "CreditLimits must be numbers but got " & CStr(CreditLimits(i, 1)) & " at position " & CStr(i)
15            If CreditLimits(i, 1) < 0 Then Throw "CreditLimits must not be negative. Got " & CStr(CreditLimits(i, 1)) & " at position " & CStr(i)
16            If CreditLimits(i, 1) > 0 Then AnyPositive = True
17        Next i
18        If Not AnyPositive Then Throw "At least one CreditLimit must be positive"

19        Exit Sub
ErrHandler:
20        Throw "#ValidateCreditInputs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HeadroomSolverFromFilters
' Author    : Philip Swannell
' Date      : 01-Sep-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function HeadroomSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
          IncludeFutureTrades As Boolean, IncludeAssetClasses As String, _
          HedgeLabels, UnitHedgeAmounts, PortfolioAgeing As Double, _
          FlipTrades As Boolean, Numeraire As String, NumberOfSims As Long, _
          TimeGap As Double, TimeEnd As Double, PFEPercentile As Double, _
          ShortfallOrQuantile As String, LimitTimes, CreditLimits, _
          CreditInterpMethod, BaseCCY As String, TradesScaleFactor As Double, _
          ByRef PFEProfileWithET, ByRef PFEProfileWithoutET, CurrenciesToInclude As String, _
          ModelName As String, TC As TradeCount, ProductCreditLimits As String, DoNotionalCap As Boolean, _
          NotionalCapApplies As Boolean, NotionalCapForNewTrades As Double, twb As Workbook, _
          ModelBareBones As Dictionary, EitherOrBasis As Boolean, ExtraTradesAre As String)

          Dim AnchorDate As Date
          Dim CalcByProduct As Boolean
          Dim Expression As String
          Dim Horizon As String
          Dim i As Long
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim IsShortfall As Boolean
          Dim JuliaFunction As String
          Dim TheTrades As Variant
          Dim UnitHedgeTrades

1         On Error GoTo ErrHandler

2         ValidateCreditInputs ShortfallOrQuantile, CreditLimits, CStr(CreditInterpMethod)
3         IsShortfall = ShortfallOrQuantile = "Shortfall"

4         AnchorDate = ModelBareBones("AnchorDate")

5         SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, True, IncludeFxTrades, IncludeRatesTrades, CalcByProduct

6         TheTrades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, FlipTrades, Numeraire, IncludeFxTrades, IncludeRatesTrades, _
              TradesScaleFactor, CurrenciesToInclude, True, TC, twb, shFutureTrades, AnchorDate)
7         UnitHedgeTrades = ConstructExtraTrades(ExtraTradesAre, ModelBareBones, UnitHedgeAmounts, HedgeLabels)

          Dim CreditLimitsLiteral As String
          Dim ExistingTradesFile As String
          Dim ExistingTradesFileX As String
          Dim LimitTimesLiteral As String
          Dim UnitHedgeTradesFile As String
          Dim UnitHedgeTradesFileX As String

8         ExistingTradesFile = LocalTemp & "HeadroomSolverExistingTrades.csv"
9         ExistingTradesFileX = MorphSlashes(ExistingTradesFile, UseLinux())
10        UnitHedgeTradesFile = LocalTemp & "HeadroomSolverUnitHedgeTrades.csv"
11        UnitHedgeTradesFileX = MorphSlashes(UnitHedgeTradesFile, UseLinux())
12        LimitTimesLiteral = "[" & sConcatenateStrings(LimitTimes) & "]"
13        CreditLimitsLiteral = "[" & sConcatenateStrings(CreditLimits) & "]"

14        ThrowIfError sCSVWrite(TheTrades, ExistingTradesFile)
15        ThrowIfError sCSVWrite(UnitHedgeTrades, UnitHedgeTradesFile)

16        If EitherOrBasis Then
17            JuliaFunction = "headroomsolvercalcbyproducteitherorbasis"
18        Else
19            JuliaFunction = "headroomsolvercalcbyproduct"
20        End If
              
21        If EitherOrBasis = False Then
22            For i = sNRows(UnitHedgeAmounts) To 1 Step -1
23                If UnitHedgeAmounts(i, 1) <> 0 Then
24                    Horizon = CStr(i)
25                    Exit For
26                End If
27            Next i
28        Else
29            Horizon = "[1"
30            For i = 2 To sNRows(UnitHedgeAmounts)
31                Horizon = Horizon + "," + CStr(i)
32            Next i
33            Horizon = Horizon + "]"
34        End If

35        Expression = "Cayley." & JuliaFunction & "(" & ModelName & ",""" & ExistingTradesFileX & """,""" & _
              UnitHedgeTradesFileX & """,""" & BaseCCY & """," & LimitTimesLiteral & "," & CreditLimitsLiteral & ",""" _
              & CreditInterpMethod & """," & CStr(NumberOfSims) & "," & CStr(TimeGap) & "," & _
              CStr(TimeEnd) & "," & CStr(PFEPercentile) & "," & LCase(IsShortfall) & "," & _
              LCase(DoNotionalCap) & "," & CStr(NotionalCapForNewTrades) & "," & LCase(CalcByProduct) & _
              "," & Horizon & "," & LCase(gUseThreads) & ")"
              
          Dim ResultFromJulia
              
36        If gDebugMode Then Debug.Print Expression

37        Assign ResultFromJulia, JuliaEvalWrapper(Expression, ModelName)

38        If VarType(ResultFromJulia) = vbString Then Throw ResultFromJulia
        
39        NotionalCapApplies = GetItem(ResultFromJulia, "notionalcapapplies")
40        PFEProfileWithoutET = GetItem(ResultFromJulia, "pfeprofilewithoutet")
41        PFEProfileWithET = GetItem(ResultFromJulia, "pfeprofilewithet")

42        HeadroomSolverFromFilters = GetItem(ResultFromJulia, "multiple")

43        Exit Function
ErrHandler:
44        HeadroomSolverFromFilters = "#HeadroomSolverFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FxSolverFromFilters
' Author    : Philip Swannell
' Date      : 28-Sep-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function FxSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
          IncludeAssetClasses As String, PortfolioAgeing As Double, FlipTrades As Boolean, _
          Numeraire As String, NumberOfSims As Long, TimeGap As Double, TimeEnd As Double, _
          PFEPercentile As Double, ShortfallOrQuantile As String, LimitTimes, CreditLimits, CreditInterpMethod As String, _
          BaseCCY As String, TradesScaleFactor As Double, ByRef PFEProfileUnshockedFx, ByRef PFEProfileShockedFx, _
          CurrenciesToInclude As String, ModelName As String, TC As TradeCount, CalcByProduct As Boolean, _
          ProductCreditLimits As String, twb As Workbook, ByRef FxRoot, ModelBareBones As Dictionary)

          Dim AnchorDate As Date
          Dim CreditLimitsLiteral As String
          Dim Expression As String
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim IsShortfall As Boolean
          Dim LimitTimesLiteral As String
          Dim ResultFromJulia
          Dim TheTrades As Variant
          Dim TradesFile As String
          Dim TradesFileX As String

1         On Error GoTo ErrHandler

2         ValidateCreditInputs ShortfallOrQuantile, CreditLimits, CreditInterpMethod

3         IsShortfall = ShortfallOrQuantile = "Shortfall"

4         AnchorDate = ModelBareBones("AnchorDate")

5         SetBooleans IncludeAssetClasses, ProductCreditLimits, "Foo", False, IncludeFxTrades, IncludeRatesTrades, CalcByProduct

6         TheTrades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, FlipTrades, Numeraire, IncludeFxTrades, IncludeRatesTrades, _
              TradesScaleFactor, CurrenciesToInclude, True, TC, twb, shFutureTrades, AnchorDate)

7         TradesFile = LocalTemp & "FxSolveTrades.csv"
8         TradesFileX = MorphSlashes(TradesFile, UseLinux())

9         ThrowIfError sCSVWrite(TheTrades, TradesFile)

10        LimitTimesLiteral = "[" & sConcatenateStrings(LimitTimes) & "]"
11        CreditLimitsLiteral = "[" & sConcatenateStrings(CreditLimits) & "]"

12        Expression = "Cayley.fxsolver(" & ModelName & ",""" & TradesFileX & """,""" & BaseCCY & _
              """," & LimitTimesLiteral & "," & CreditLimitsLiteral & ", """ & CreditInterpMethod & """, " & _
              CStr(NumberOfSims) & ", " & CStr(TimeGap) & ", " & CStr(TimeEnd) & ", " & _
              CStr(PFEPercentile) & "," & LCase(CStr(IsShortfall)) & "," & LCase(CalcByProduct) & _
              "," & LCase(gUseThreads) & ")"

13        If gDebugMode Then Debug.Print Expression

14        Assign ResultFromJulia, JuliaEvalWrapper(Expression, ModelName)

15        If VarType(ResultFromJulia) = vbString Then Throw ResultFromJulia
        
16        PFEProfileUnshockedFx = GetItem(ResultFromJulia, "pfeprofileunshockedfx")
17        PFEProfileShockedFx = GetItem(ResultFromJulia, "pfeprofileshockedfx")
18        FxSolverFromFilters = GetItem(ResultFromJulia, "multiple")
19        FxRoot = GetItem(ResultFromJulia, "fxroot")

20        Exit Function
ErrHandler:
21        FxSolverFromFilters = "#FxSolverFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
27    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FxVolSolverFromFilters
' Author     : Philip Swannell
' Date       : 26-Jan-2022
' Purpose    : Wraps Julia function fxvolsolver
' -----------------------------------------------------------------------------------------------------------------------
Function FxVolSolverFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades As Boolean, _
          IncludeAssetClasses As String, PortfolioAgeing As Double, FlipTrades As Boolean, _
          Numeraire As String, NumberOfSims As Long, TimeGap As Double, TimeEnd As Double, _
          PFEPercentile As Double, ShortfallOrQuantile As String, LimitTimes, CreditLimits, CreditInterpMethod As String, _
          BaseCCY As String, TradesScaleFactor As Double, ByRef PFEProfileUnshockedFx, ByRef PFEProfileShockedFx, _
          CurrenciesToInclude As String, ModelName As String, TC As TradeCount, CalcByProduct As Boolean, _
          ProductCreditLimits As String, twb As Workbook, ByRef FxVolRoot, ModelBareBones As Dictionary)

          Dim AnchorDate As Date
          Dim CreditLimitsLiteral As String
          Dim Expression As String
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim IsShortfall As Boolean
          Dim LimitTimesLiteral As String
          Dim ResultFromJulia
          Dim TheTrades As Variant
          Dim TradesFile As String
          Dim TradesFileX As String

1         On Error GoTo ErrHandler

2         ValidateCreditInputs ShortfallOrQuantile, CreditLimits, CreditInterpMethod
3         IsShortfall = ShortfallOrQuantile = "Shortfall"
4         SetBooleans IncludeAssetClasses, ProductCreditLimits, "Foo", False, IncludeFxTrades, IncludeRatesTrades, CalcByProduct
5         AnchorDate = ModelBareBones("AnchorDate")

6         TheTrades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, FlipTrades, Numeraire, IncludeFxTrades, IncludeRatesTrades, _
              TradesScaleFactor, CurrenciesToInclude, True, TC, twb, shFutureTrades, AnchorDate)

7         TradesFile = LocalTemp & "FxVolSolveTrades.csv"
8         TradesFileX = MorphSlashes(TradesFile, UseLinux())

9         ThrowIfError sCSVWrite(TheTrades, TradesFile)

10        LimitTimesLiteral = "[" & sConcatenateStrings(LimitTimes) & "]"
11        CreditLimitsLiteral = "[" & sConcatenateStrings(CreditLimits) & "]"

12        Expression = "Cayley.fxvolsolver(" & ModelName & ",""" & TradesFileX & """,""" & BaseCCY & _
              """," & LimitTimesLiteral & "," & CreditLimitsLiteral & ", """ & CreditInterpMethod & """, " & _
              CStr(NumberOfSims) & ", " & CStr(TimeGap) & ", " & CStr(TimeEnd) & ", " & _
              CStr(PFEPercentile) & "," & LCase(CStr(IsShortfall)) & "," & LCase(CalcByProduct) & _
              "," & LCase(gUseThreads) & ")"

13        If gDebugMode Then Debug.Print Expression

14        Assign ResultFromJulia, JuliaEvalWrapper(Expression, ModelName)

15        If VarType(ResultFromJulia) = vbString Then Throw ResultFromJulia
        
16        PFEProfileUnshockedFx = GetItem(ResultFromJulia, "pfeprofileunshockedfx")
17        PFEProfileShockedFx = GetItem(ResultFromJulia, "pfeprofileshockedfx")
18        FxVolRoot = GetItem(ResultFromJulia, "fxvolroot")
19        FxVolSolverFromFilters = GetItem(ResultFromJulia, "multiple")

20        Exit Function
ErrHandler:
21        FxVolSolverFromFilters = "#FxVolSolverFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
27    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PFEProfileHW
' Author    : Philip Swannell
' Date      : 14-Jul-2016
' Purpose   : Wraps Julia function pfeprofilefortrades
' -----------------------------------------------------------------------------------------------------------------------
Function PFEProfileHW(Model As String, NumberOfSims As Long, ByVal TheTrades As Variant, TheExtraTrades, _
          TimeGap As Double, TimeEnd As Double, PFEPercentile As Double, IsShortfall As Boolean, _
          ReportCurrency As String, CalcByProduct As Boolean)
          
          Dim Expression As String
          Dim ExtraTradesFile As String
          Dim ExtraTradesFileX As String
          Dim TradeFilesForJulia As String
          Dim TradesFile As String
          Dim TradesFileX As String
          Dim ul As Boolean

1         On Error GoTo ErrHandler
2         If TypeName(TheTrades) = "Range" Then TheTrades = TheTrades.Value2

3         TradesFile = LocalTemp & "PFECalcTrades.csv"
4         ExtraTradesFile = LocalTemp & "PFECalcExtraTrades.csv"
5         ul = UseLinux()
6         TradesFileX = MorphSlashes(TradesFile, ul)
7         ExtraTradesFileX = MorphSlashes(ExtraTradesFile, ul)
8         ThrowIfError sCSVWrite(TheTrades, TradesFile)

9         If sNRows(TheExtraTrades) > 1 Then
10            ThrowIfError sCSVWrite(TheExtraTrades, ExtraTradesFile)
              'Pass the names of two trades files as a Julia literal array of two strings
11            TradeFilesForJulia = "[""" & TradesFileX & """,""" & ExtraTradesFileX & """]"
12        Else
13            TradeFilesForJulia = """" & TradesFileX & """"
14        End If

15        Expression = "Cayley.pfeprofilefortrades(" & Model & "," & NumberOfSims & "," & TradeFilesForJulia & "," & _
              CStr(TimeGap) & "," & CStr(TimeEnd) & "," & CStr(PFEPercentile) & "," & LCase(CStr(IsShortfall)) & _
              "," & LCase(CStr(gUseThreads)) & ",""" & UCase(ReportCurrency) & """," & LCase(CalcByProduct) & ")"

16        If gDebugMode Then Debug.Print Expression
17        PFEProfileHW = JuliaEvalWrapper(Expression, Model)
18        Exit Function
ErrHandler:
19        PFEProfileHW = "#PFEProfileHW (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PFEProfileFromFiltersPCL
' Author    : Philip Swannell
' Date      : 02-Sep-2016
' Purpose   : Wrapper to PFEProfileFromFilters, but handling the three ProductCreditLimits cases
' -----------------------------------------------------------------------------------------------------------------------
Function PFEProfileFromFiltersPCL(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
          IncludeFutureTrades As Boolean, ByVal IncludeAssetClasses As String, ByVal IncludeExtraTrades As Boolean, _
          ExtraTradeLabels, ExtraTradeAmounts, PortfolioAgeing As Double, FlipTrades As Boolean, _
          TradesScaleFactor As Double, BaseCCY As String, Model As String, NumSims As Long, TimeGap As Double, _
          TimeEnd As Double, Methodology As String, ByVal PFEPercentile, ShortfallOrQuantile As String, _
          FxNotionalPercentages, RatesNotionalPercentages, CurrenciesToInclude As String, ModelName As String, _
          TC As TradeCount, ProductCreditLimits As String, twb As Workbook, CompressTrades As Boolean, _
          ModelBareBones As Dictionary, ExtraTradesAre As String)

1         On Error GoTo ErrHandler
          Dim CalcByProduct As Boolean
          Dim i As Long
          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim Res
          Dim Res1
          Dim res2
          Dim TC1 As TradeCount
          Dim TC2 As TradeCount
          
2         SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, IncludeExtraTrades, IncludeFxTrades, IncludeRatesTrades, CalcByProduct

3         If CalcByProduct Or LCase(Methodology) <> "notional based" Then
4             PFEProfileFromFiltersPCL = PFEProfileFromFilters(FilterBy1, Filter1Value, FilterBy2, _
                  Filter2Value, IncludeFutureTrades, IncludeAssetClasses, IncludeExtraTrades, ExtraTradeLabels, _
                  ExtraTradeAmounts, PortfolioAgeing, FlipTrades, TradesScaleFactor, BaseCCY, Model, NumSims, _
                  TimeGap, TimeEnd, Methodology, PFEPercentile, ShortfallOrQuantile, FxNotionalPercentages, _
                  RatesNotionalPercentages, CurrenciesToInclude, ModelName, TC, ProductCreditLimits, twb, _
                  CompressTrades, ModelBareBones, CalcByProduct, ExtraTradesAre)
5         Else

6             If IncludeFxTrades Then
7                 Res1 = PFEProfileFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
                      IncludeFutureTrades, "Fx", IncludeExtraTrades, ExtraTradeLabels, ExtraTradeAmounts, _
                      PortfolioAgeing, FlipTrades, TradesScaleFactor, BaseCCY, Model, NumSims, _
                      TimeGap, TimeEnd, Methodology, PFEPercentile, ShortfallOrQuantile, FxNotionalPercentages, _
                      RatesNotionalPercentages, CurrenciesToInclude, ModelName, TC1, ProductCreditLimits, twb, _
                      CompressTrades, ModelBareBones, CalcByProduct, ExtraTradesAre)
8             End If
9             If IncludeRatesTrades Then
                  'Must not include extra trades twice
10                If IncludeFxTrades Then IncludeExtraTrades = False
11                res2 = PFEProfileFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
                      IncludeFutureTrades, "Rates", IncludeExtraTrades, ExtraTradeLabels, ExtraTradeAmounts, _
                      PortfolioAgeing, FlipTrades, TradesScaleFactor, BaseCCY, Model, NumSims, _
                      TimeGap, TimeEnd, Methodology, PFEPercentile, ShortfallOrQuantile, FxNotionalPercentages, _
                      RatesNotionalPercentages, CurrenciesToInclude, ModelName, TC2, ProductCreditLimits, twb, _
                      CompressTrades, ModelBareBones, CalcByProduct, ExtraTradesAre)
12            End If
13            If IncludeRatesTrades And IncludeFxTrades Then
14                Res = Res1
15                If LBound(Res1) <> LBound(res2) Or UBound(Res1) <> UBound(res2) Or _
                      LBound(Res1, 2) <> LBound(res2, 2) Or UBound(Res1, 2) <> UBound(res2, 2) Then
16                    Throw "Assertion Failed, mismatch in array size in calls to PFEProfileFromFilters"
17                End If
                  Dim FirstColNo As Long
18                FirstColNo = LBound(Res1, 2)
19                For i = LBound(Res1) To UBound(Res1)
20                    If Res1(i, FirstColNo) <> res2(i, FirstColNo) Then
21                        Throw "Assertion Failed: Date mismatch in calls to PFEProfileFromFilters"
22                    End If
23                    If Res1(i, FirstColNo + 1) <> res2(i, FirstColNo + 1) Then
24                        Throw "Assertion Failed: Time mismatch in calls to PFEProfileFromFilters"
25                    End If
26                    Res(i, FirstColNo + 2) = Res1(i, FirstColNo + 2) + res2(i, FirstColNo + 2)
27                Next i
28                PFEProfileFromFiltersPCL = Res
29                TC.NumExcluded = TC1.NumExcluded + TC2.NumExcluded
30                TC.NumIncluded = TC1.NumIncluded + TC2.NumIncluded
31                TC.Total = TC1.Total + TC2.Total
32            ElseIf IncludeRatesTrades Then
33                PFEProfileFromFiltersPCL = res2
34                TC = TC2
35            Else
36                PFEProfileFromFiltersPCL = Res1
37                TC = TC1
38            End If
39        End If
40        Exit Function
ErrHandler:
41        PFEProfileFromFiltersPCL = "#PFEProfileFromFiltersPCL (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PFEProfileFromFilters
' Author    : Philip Swannell
' Date      : 02-Sep-2016
' Purpose   : Wrapper to one of PFEProfileHWFromFilters, PFEProfileLnFxFromFilters, NotionalBasedByFilters
' -----------------------------------------------------------------------------------------------------------------------
Function PFEProfileFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
          IncludeFutureTrades As Boolean, IncludeAssetClasses As String, IncludeExtraTrades As Boolean, _
          ExtraTradeLabels, ExtraTradeAmounts, PortfolioAgeing As Double, FlipTrades As Boolean, _
          TradesScaleFactor As Double, BaseCCY As String, Model As String, NumSims As Long, TimeGap As Double, _
          TimeEnd As Double, Methodology As String, ByVal PFEPercentile, ShortfallOrQuantile As String, _
          FxNotionalPercentages, RatesNotionalPercentages, CurrenciesToInclude As String, ModelName As String, _
          TC As TradeCount, ProductCreditLimits As String, twb As Workbook, CompressTrades As Boolean, _
          ModelBareBones As Dictionary, CalcByProduct As Boolean, ExtraTradesAre As String)

          Dim IncludeFxTrades As Boolean
          Dim IncludeRatesTrades As Boolean
          Dim IsShortfall As Boolean
          Dim Numeraire As String

1         On Error GoTo ErrHandler
2         Numeraire = GetItem(ModelBareBones, "Numeraire")

3         SetBooleans IncludeAssetClasses, ProductCreditLimits, ExtraTradesAre, IncludeExtraTrades, IncludeFxTrades, IncludeRatesTrades, CalcByProduct

4         If LCase(Methodology) = "notional based" Then
5             PFEProfileFromFilters = ThrowIfError(NotionalBasedByFilters(FilterBy1, Filter1Value, FilterBy2, _
                  Filter2Value, BaseCCY, FxNotionalPercentages, RatesNotionalPercentages, IncludeExtraTrades, _
                  IncludeAssetClasses, ExtraTradeLabels, ExtraTradeAmounts, IncludeFutureTrades, PortfolioAgeing, _
                  TradesScaleFactor, CurrenciesToInclude, ModelName, TC, TimeEnd, TimeGap, ProductCreditLimits, twb, _
                  shFutureTrades, ModelBareBones, ExtraTradesAre))
6         Else
7             If LCase(ShortfallOrQuantile) = "shortfall" Then
8                 IsShortfall = True
9             ElseIf LCase(ShortfallOrQuantile) = "quantile" Then
10                IsShortfall = False
11            Else
12                Throw "ShortfallOrQuantile must be either 'Shortfall' or 'Quantile'"
13            End If

14            PFEProfileFromFilters = ThrowIfError(PFEProfileHWFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeFutureTrades, IncludeFxTrades, _
                  IncludeRatesTrades, IncludeExtraTrades, ExtraTradeLabels, ExtraTradeAmounts, PortfolioAgeing, _
                  FlipTrades, Numeraire, Model, NumSims, TimeGap, TimeEnd, CDbl(PFEPercentile), IsShortfall, _
                  BaseCCY, CDbl(TradesScaleFactor), CurrenciesToInclude, TC, twb, CompressTrades, ModelBareBones, CalcByProduct, ExtraTradesAre))
15        End If

16        Exit Function
ErrHandler:
17        Throw "#PFEProfileFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeFromMarketDataBook
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : We want to get values from the Market Data workbook without setting up the
'             dreaded Excel links.
'             But note that we should be accessing market data by interrogating the
'             R environment - see functions in module modMarketDataFromR.
' -----------------------------------------------------------------------------------------------------------------------
Function RangeFromMarketDataBook(SheetName As String, RangeName As String)
1         Application.Volatile
          Dim BookName As String
          Dim ErrString As String
          Dim N As Name
          Dim wb As Workbook
          Dim ws As Worksheet

2         On Error GoTo ErrHandler
3         BookName = RangeFromSheet(shConfig, "MarketDataWorkbook")
          'Can't use sSplitPath here, since may be only a filename
4         If InStr(BookName, "\") > 0 Then
5             BookName = Mid(BookName, InStrRev(BookName, "\") + 1)
6         End If

7         On Error Resume Next
8         Set RangeFromMarketDataBook = Application.Workbooks(BookName).Worksheets(SheetName).Names(RangeName).RefersToRange
9         If Err.Number = 0 Then Exit Function
10        On Error GoTo ErrHandler

11        If Not IsInCollection(Application.Workbooks, BookName) Then Throw "Market Data Workbook (" & BookName & ") is not open"
12        Set wb = Application.Workbooks(BookName)
13        If Not IsInCollection(wb.Worksheets, SheetName) Then Throw "#Cannot find worksheet '" & SheetName & "' in MarketDataWorkbook"
14        Set ws = wb.Worksheets(SheetName)
15        If Not IsInCollection(ws.Names, RangeName) Then Throw "#Cannot find range named '" & RangeName & "' in sheet '" & SheetName & "' of MarketDataWorkbook"
16        Set N = ws.Names(RangeName)
17        If Not NameRefersToRange(N) Then Throw "Name '" & RangeName & "' on sheet '" & SheetName & "' of Market Data book does not refer to a range"
18        Throw "Unknown Error"

19        Exit Function
ErrHandler:
20        ErrString = "#RangeFromMarketDataBook (line " & CStr(Erl) & "): " & Err.Description & "!"
21        If TypeName(Application.Caller) = "Range" Then
22            RangeFromMarketDataBook = ErrString
23        Else
24            Throw ErrString
25        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NameRefersToRange
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Returns TRUE if the name refers to a Range, FALSE if it refers to something else.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NameRefersToRange(TheName As Name) As Boolean
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = TheName.RefersToRange
3         NameRefersToRange = True
4         Exit Function
ErrHandler:
5         NameRefersToRange = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SolveFxHeadroom
' Author    : Philip Swannell
' Date      : 25-May-2015
' Purpose   : Solves for the Fx headroom (multiplicative shock to EUR against all currencies)
'             via roll-my-own solver and successive running of the PFE sheet in "Standard" mode.
'             Now (Jan 2022) only used for Notional-based.
' -----------------------------------------------------------------------------------------------------------------------
Function SolveFxHeadroom(ThrowErrors As Boolean, DisplayBaseCasePFE As Boolean, ByRef Success As Boolean)
          Dim CalcCount As Long
          Dim FxShockCell As Range
          Dim HeadroomCell As Range
          Dim x1 As Double
          Dim x2 As Double
          Dim x3 As Double
          Dim y1 As Double
          Dim y2 As Double
          Dim y3 As Double
          Const MaxNumCalcs = 15
          Const Tolerance = 1000
          Dim Filter1Value As Variant
          Dim HH As Long
          Dim IsNB As Boolean
          Dim Methodology As String
          Dim ModelBareBones As Dictionary
          Dim origFxShock As Variant
          Dim PreviousSolution As Double
          Dim SPH As Object
          Dim UseHistoricalFxVol As Boolean
          Dim VolatilityInput As String

1         On Error GoTo ErrHandler
        
2         Filter1Value = RangeFromSheet(shCreditUsage, "Filter1Value").Value
3         Methodology = FirstElement(LookupCounterpartyInfo(Filter1Value, "Methodology", " - ", " - "))
4         IsNB = LCase(Methodology) = "notional based"
5         VolatilityInput = IIf(IsNB, "-", FirstElement(sIfErrorString(LookupCounterpartyInfo(Filter1Value, _
              "Volatility Input"), "MARKET IMPLIED")))
6         UseHistoricalFxVol = VolatilityInput = "HISTORICAL"
7         Set ModelBareBones = IIf(UseHistoricalFxVol, gModel_CMHS, gModel_CMS)

8         origFxShock = RangeFromSheet(shCreditUsage, "FxShock")
9         Set SPH = CreateSheetProtectionHandler(shCreditUsage)
10        RangeFromSheet(shCreditUsage, "FxHeadroom").ClearContents

11        HH = GetHedgeHorizon()

12        Set HeadroomCell = RangeFromSheet(shCreditUsage, "MinHeadroomOverFirstN").Cells(HH, 1)
13        Set FxShockCell = RangeFromSheet(shCreditUsage, "FxShock")
14        RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value2 = False

15        x1 = 1
16        x2 = 1.01

17        FxShockCell.Value = x1
18        RunCreditUsageSheet "Standard", True, False, True
19        If Not IsNumber(HeadroomCell.Value) Then Throw "Non-numbers in Headroom range"
20        y1 = HeadroomCell.Value

21        FxShockCell.Value = x2
22        RunCreditUsageSheet "Standard", True, False, True: CalcCount = CalcCount + 1
23        y2 = HeadroomCell.Value

TryAgain:
24        If Not IsNumber(HeadroomCell.Value) Then GoTo NoSolutionFound
25        y2 = HeadroomCell.Value
26        x3 = x1 - y1 * (x2 - x1) / (y2 - y1)
27        If x3 < 0.001 Then x3 = 0.001        'since negative FxShock will certainly cause errors

28        FxShockCell.Value = x3
29        RunCreditUsageSheet "Standard", True, False, True: CalcCount = CalcCount + 1
30        If Not IsNumber(HeadroomCell.Value) Then GoTo NoSolutionFound
31        y3 = HeadroomCell.Value
32        If Not IsNumber(HeadroomCell.Value) Then GoTo NoSolutionFound
33        If Abs(y3) > Tolerance And CalcCount < MaxNumCalcs Then
34            x1 = x2: x2 = x3: y1 = y2: y2 = y3
35            GoTo TryAgain
36        End If

FoundSolution:
37        If CalcCount < MaxNumCalcs Then
38            Success = True
39            RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 1).Value2 = _
                  RangeFromSheet(shCreditUsage, "EURUSD")
40            RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 2).Value2 = x3
41            PreviousSolution = x3
42        Else
NoSolutionFound:
43            Success = False
44            RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 1).Value2 = "#No solution found!"
45            RangeFromSheet(shCreditUsage, "FxHeadroom").Cells(1, 2).Value2 = "#No solution found!"
46        End If

47        CalcCount = 0

48        FxShockCell.Value = origFxShock
49        If DisplayBaseCasePFE Then
50            RunCreditUsageSheet "Standard", True, False, True
51        End If
52        SolveFxHeadroom = "OK"
53        RangeFromSheet(shCreditUsage, "FxSolveResult").Value = "OK"
54        Exit Function
ErrHandler:

55        SolveFxHeadroom = "#SolveFxHeadroom (line " & CStr(Erl) & "): " & Err.Description & "!"
56        RangeFromSheet(shCreditUsage, "FxSolveResult").Value = SolveFxHeadroom

57        If Not FxShockCell Is Nothing Then
58            FxShockCell.Value = origFxShock
59        End If
60        If ThrowErrors Then Throw SolveFxHeadroom
End Function

