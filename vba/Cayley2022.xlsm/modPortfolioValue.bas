Attribute VB_Name = "modPortfolioValue"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PortfolioValueFromFilters
' Author    : Philip Swannell
' Date      : 13-Jul-2015
' Purpose   : Returns the value of the trades brought back by the passed in filters, with portfolio ageing applied
'             valuation uses the shocks currently entered on the CreditUsage sheet. Function assumes that
'             the market data sheets, FxVols and DiscountFactors, have been recalculated.
' -----------------------------------------------------------------------------------------------------------------------
Function PortfolioValueFromFilters(BaseCCY As String, FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
          IncludeFutureTrades As Boolean, IncludeAssetClasses As String, PortfolioAgeing As Double, _
          TradesScaleFactor As Double, CurrenciesToInclude As String, ModelName As String, _
          TC As TradeCount, ProductCreditLimits As String, twb As Workbook, AnchorDate As Date)
          
          Dim FlipTrades As Boolean
          Dim Numeraire As String
          Dim NumTrades As Long
          Dim TheTrades As Variant
          Dim WithFxTrades As Boolean
          Dim WithRatesTrades As Boolean

1         On Error GoTo ErrHandler

2         FlipTrades = True

3         Numeraire = NumeraireFromMDWB()
4         SetBooleans IncludeAssetClasses, ProductCreditLimits, "Foo", False, WithFxTrades, WithRatesTrades, False

5         TheTrades = GetTradesInJuliaFormat(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
              IncludeFutureTrades, PortfolioAgeing, _
              FlipTrades, Numeraire, WithFxTrades, WithRatesTrades, TradesScaleFactor, _
              CurrenciesToInclude, True, TC, twb, shFutureTrades, AnchorDate)
6         NumTrades = TC.NumIncluded
7         If NumTrades = 0 Then
8             PortfolioValueFromFilters = 0
9         Else
10            PortfolioValueFromFilters = PortfolioValueHW(TheTrades, ModelName, BaseCCY)
11        End If

12        Exit Function
ErrHandler:
13        PortfolioValueFromFilters = "#PortfolioValueFromFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PortfolioValueHW
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : Wrap Julia function valueportfolio. Returns a vector of trade values with embedded error strings where
'             trades are not possible to value.
' -----------------------------------------------------------------------------------------------------------------------
Function PortfolioValueHW(TheTrades, ModelName, ReportCurrency, Optional ReturnVector As Boolean = False)
          Dim Expression As String
          Dim ThrowOnError As Boolean
          Dim TradeFile As String
          Dim TradeFileX As String

1         On Error GoTo ErrHandler

2         ThrowOnError = Not (ReturnVector)
3         TradeFile = LocalTemp() & "CayleyTrades2.csv"
4         TradeFileX = MorphSlashes(TradeFile, UseLinux())

5         ThrowIfError sCSVWrite(TheTrades, TradeFile)
6         Expression = "Cayley.valueportfolio(" & ModelName & ",""" & TradeFileX & """, 0.0,""" & _
              ReportCurrency & """," & LCase(ThrowOnError) & "," & LCase(ReturnVector) & "," & LCase(gUseThreads) & ")"
              
7         If gDebugMode Then Debug.Print (Expression)
8         PortfolioValueHW = JuliaEvalVBA(Expression)
9         If VarType(PortfolioValueHW) = vbString Then Throw PortfolioValueHW

10        Exit Function
ErrHandler:
11        Throw "#PortfolioValueHW (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FirstElement
' Author    : Philip Swannell
' Date      : 06-Jul-2015
' Purpose   : Assumes Data is either a 2-d array or not an array.
'             Throws if the return would be an error
' -----------------------------------------------------------------------------------------------------------------------
Function FirstElement(Data)
1         On Error GoTo ErrHandler
2         If VarType(Data) < vbArray Then
3             If VarType(Data) = vbString Then If Left(Data, 1) = "#" Then If Right(Data, 1) = "!" Then Throw CStr(Data)
4             FirstElement = Data
5         Else
6             FirstElement = Data(1, 1)
7             If VarType(FirstElement) = vbString Then
8                 If Left(FirstElement, 1) = "#" Then
9                     If Right(FirstElement, 1) = "!" Then
10                        Throw CStr(FirstElement)
11                    End If
12                End If
13            End If
14        End If
15        Exit Function
ErrHandler:
16        Throw "#FirstElement (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
