Attribute VB_Name = "modFxOptionPFE"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketOptionCapFloor
' Author    : Philip Swannell
' Date      : 16-Jan-2019
' Purpose   : Where S is the (weighted) mean of N correlated log-normal assets, returns the undiscounted
'             value of an option that pays K - S (for a Put) or S - K (for a Call), subject
'             to a Cap and Floor on the payout. Valuation by Monte Carlo.
' Arguments
' Forwards  : A vector of the forward prices of each asset (can be 1-row or 1-column).
' Vols      : A vector of the log-normal volatilities of each asset (can be 1-row or 1-column).
' Correlations: The correlation matrix. Must be positive definite with on-diagonal elements of 1.
' Time      : Time in years until exercise.
' NumPaths  : The number of Monte Carlo paths. If RNG is "Sobol" then NumPaths must be one less than a
'             power of 2.
' RNGName   : The name of the Random Number Generator, allowed values are "Sobol", "Mersenne-Twister" or
'             "Wichmann-Hill"
' Notional  : The option Notional.
' Strike    : The option Strike
' Cap       : The option Cap, the payout is never more than this.
' Floor     : The option Floor, the payout is never less than this.
' CallOrPut : Determines if the option is a Call or Put. For a Call enter TRUE, "C" or "CALL". For a Put
'             enter FALSE,"P" or "PUT".
' Weights   : The weights used to calculated the weighted mean in the payoff definition. Must be a 1-row
'             or 1-column array of the same size as Forwards. Omit for unweighted mean
'             (equivalent to a vector of length N, with elements 1/N).
'
' Notes     : If S is the arithmetic (weighted) mean of N correlated assets, and K is the strike:
'             For a Call, the option pays:
'             Notional x Middle((S - K, Cap, Floor)
'             For a Put the option pays:
'             Notional x Middle((K - S, Cap, Floor)
'             where Middle(A,B,C) is the middle value of A, B and C, Middle(A,B,C) =
'             Min(Max(A,B),Max(A,C),Max(B,C))
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketOptionCapFloor(Forwards As Variant, Vols As Variant, Correlations As Variant, Time As Double, NumPaths As Long, RNGName As String, Notional As Double, Strike As Double, Cap As Double, Floor As Double, CallOrPut As Variant, Optional Weights)
Attribute sBasketOptionCapFloor.VB_Description = "Where S is the (weighted) mean of N correlated log-normal assets, returns the undiscounted value of an option that pays K - S (for a Put) or S - K (for a Call), subject to a Cap and Floor on the payout. Valuation by Monte Carlo."
Attribute sBasketOptionCapFloor.VB_ProcData.VB_Invoke_Func = " \n29"

1         On Error GoTo ErrHandler
          Dim CappedPayoff
          Dim i As Long
          Dim IsCall As Boolean
          Dim Paths
          Dim TakeMean As Boolean
          Dim UnCappedPayoff
          
2         If VarType(CallOrPut) = vbBoolean Then
3             IsCall = CallOrPut
4         ElseIf VarType(CallOrPut) = vbString Then
5             Select Case UCase$(CallOrPut)
                  Case "C", "CALL"
6                     IsCall = True
7                 Case "P", "PUT"
8                     IsCall = False
9                 Case Else
10                    Throw "CallOrPut must be TRUE, 'C' or 'CALL' for a Call, or FALSE, 'P' or 'PUT' for a Put"
11            End Select
12        Else
13            Throw "CallOrPut must be TRUE, 'C' or 'CALL' for a Call, or FALSE, 'P' or 'PUT' for a Put"
14        End If
          
          'Ensure Forwards is a Column vector
15        If sNRows(Forwards) = 1 Then
16            If sNCols(Forwards) > 1 Then
17                Forwards = sArrayTranspose(Forwards)
18            End If
19        End If
          
          'Ensure Vols is a Column vector
20        If sNRows(Vols) = 1 Then
21            If sNCols(Vols) > 1 Then
22                Vols = sArrayTranspose(Vols)
23            End If
24        End If
          
          'Ensure Weights is a ROW vector
25        If IsMissing(Weights) Then
26            TakeMean = True
27        Else
28            Force2DArrayR Weights
29            TakeMean = False
30            If sNCols(Weights) = 1 Then
31                If sNRows(Vols) > 1 Then
32                    Weights = sArrayTranspose(Weights)
33                End If
34            End If
35            For i = 1 To sNCols(Weights)
36                If Not IsNumber(Weights(1, i)) Then Throw "Weights must be numbers"
37            Next i
38        End If
          
39        If sNRows(Forwards) <> sNRows(Vols) Then
40            Throw "Forwards, Vols and Weights (if provided) must all be 1-row or 1-column arrays and all have the same number of elements"
41            If Not TakeMean Then
42                If sNRows(Forwards) <> sNCols(Weights) Then
43                    Throw "Forwards, Vols and Weights (if provided) must all be 1-row or 1-column arrays and all have the same number of elements"
44                End If
45            End If
46        End If
          
47        If Cap < Floor Then Throw "Cap must be greater than or equal to Floor"
48        Paths = ThrowIfError(sCorrelatedLogNormals(Forwards, Vols, Correlations, Time, NumPaths, RNGName))
49        If IsCall Then
50            If TakeMean Then
51                UnCappedPayoff = sArraySubtract(sRowMean(Paths), Strike)
52            Else
53                UnCappedPayoff = sArraySubtract(sRowSum(sArrayMultiply(Paths, Weights)), Strike)
54            End If
55        Else
56            If TakeMean Then
57                UnCappedPayoff = sArraySubtract(Strike, sRowMean(Paths))
58            Else
59                UnCappedPayoff = sArraySubtract(Strike, sRowSum(sArrayMultiply(Paths, Weights)))
60            End If
61        End If

62        CappedPayoff = sArrayMin(sArrayMax(UnCappedPayoff, Floor), Cap)
63        sBasketOptionCapFloor = sColumnMean(CappedPayoff)(1, 1) * Notional
          
64        Exit Function
ErrHandler:
65        sBasketOptionCapFloor = "#sBasketOptionCapFloor (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCorrelatedLogNormals
' Author    : Philip Swannell
' Date      : 01-Jul-2015
' Purpose   : Returns a sample of size NumPaths from a multi-variate log-normal distribution. Except
'             when RNGName = "Sobol" the sample is adjusted to precisely match the input
'             forwards and vols.
' Arguments
' Forwards  : Enter a column vector giving the mean of the multi-variate distribution.
' Vols      : Enter a column vector giving the volatilities (standard deviations divided by square root
'             of time) of the underlying normal distributions
' Correlations: The correlation matrix. Must be positive definite with on-diagonal elements of 1.
' Time      : The time horizon, i.e. the standard deviations of the underlying normal distributions are
'             given by squareroot(Time) x Vols
' NumPaths  : The sample size. The returned array will have this many rows, and as many columns as there
'             are rows in Forwards.
' RNGName   : The name of the Random Number Generator, allowed values are "Sobol", "Mersenne-Twister" or
'             "Wichmann-Hill"
' -----------------------------------------------------------------------------------------------------------------------
Function sCorrelatedLogNormals(ByVal Forwards, ByVal Vols, ByVal Correlations, Time As Double, NumPaths As Long, RNGName As String, Optional Seed As Variant)
Attribute sCorrelatedLogNormals.VB_Description = "Returns a sample of size NumPaths from a multi-variate log-normal distribution. Except when RNGName = ""Sobol"" the sample is adjusted to precisely match the input forwards and vols."
Attribute sCorrelatedLogNormals.VB_ProcData.VB_Invoke_Func = " \n23"

1         On Error GoTo ErrHandler

          Dim Cholesky As Variant
          Dim CorrelatedNormals
          Dim CurrentMeans As Variant
          Dim CurrentStDev As Variant
          Dim DesiredMeans As Variant
          Dim DesiredStDev As Variant
          Dim Normals As Variant
          Dim NumAssets As Long
          Dim Rebase As Boolean
          Dim RebasedCorrelatedNormals

2         Force2DArrayRMulti Forwards, Vols, Correlations

3         If NumPaths < 1 Then Throw "NumPaths must be positive"
4         NumAssets = sNRows(Forwards)

5         Normals = ThrowIfError(sRandomVariable(NumPaths, NumAssets, "Normal", RNGName, , Seed))
6         If sArraysIdentical(Correlations, sIdentityMatrix(sNRows(Forwards))) Then
7             CorrelatedNormals = Normals
8         Else
9             Cholesky = ThrowIfError(sCholesky(Correlations))
10            CorrelatedNormals = Application.WorksheetFunction.MMult(Normals, sArrayTranspose(Cholesky))
11        End If

12        Rebase = InStr(LCase$(RNGName), "sobol") = 0        'We don't correct mean or standard dev for sobol sequences

13        DesiredMeans = sArrayTranspose(sArraySubtract(sArrayLog(Forwards), sArrayMultiply(0.5, Vols, Vols, Time)))
14        DesiredStDev = sArrayTranspose(sArrayMultiply(Vols, sArrayPower(Time, 0.5)))
15        If Rebase Then
16            CurrentMeans = sColumnMean(CorrelatedNormals)
17            CurrentStDev = sColumnStDev(CorrelatedNormals)
18            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(sArraySubtract(CorrelatedNormals, CurrentMeans), sArrayDivide(DesiredStDev, CurrentStDev)), DesiredMeans)
19        Else
20            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(CorrelatedNormals, DesiredStDev), DesiredMeans)
21        End If
22        sCorrelatedLogNormals = sArrayExp(RebasedCorrelatedNormals)

23        Exit Function
ErrHandler:
24        sCorrelatedLogNormals = "#sCorrelatedLogNormals (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFxOptionPFE
' Author    : Philip Swannell
' Date      : 02-Jul-2015
' Purpose   : Calculates the PFE on a portfolio of FxOptions
' Arguments
' AnchorDate: The "Today Date" for calculation of PFE.
' HorizonDate: The date to which PFE is calculated.
' BaseCcy   : String indicating currency in which the PFE is returned.
' NumPaths  : Number of Monte-Carlo or Sobol paths
' RNGName   : The name of the Random Number Generator, allowed values are "Sobol", "Mersenne-Twister" or
'             "Wichmann-Hill"
' Confidence: Confidence level, must be between 1/NumPaths and 1 - 1/NumPaths and can be a column array
'             (in which case the return from the function is a volumn array). 0.95 for 95%
'             confidence level.
' IsShortfall: IsShortfall. If TRUE, then the return is the conditional expectation of the portfolio
'             value, conditional on the value exceeding the confidence level percentile. If
'             FALSE, then the return is the confidence level percentile.
' BuySell   : If the option is a long position then BuySell should be "B", otherwise it should be "S".
'             This argument and the next 13 arguments can be column arrays (of the same
'             size) to handle portfolios of Fx options.
' OptionStyle: can take values: "C" or "Call", "P" or "Put", "B" or "Buy", "S" or "Sell", "UD" or "Up
'             Digital", "DD" or "Down Digital". Case insensitive and space characters are
'             ignored so "Up Digital" is equivalent to "updigital".
' CCY1      : The currency code of the first currency. The option is an option on this currency versus
'             CCY2.
' CCY2      : The currency code of the second currency. The payout of the option is denominated in this
'             currency.
' Amount1   : The amount of CCY1 on which the option is written. Must be positive when BuySell is B,
'             negaitive when BuySell is S
' Amount2   : Must be Amount1 * -Strike (or else an error is returned)
' Expiry    : The expiry date of the option
' Strike    : The Strike of the option.
' VolToHrzn : The volatility of the CCy1/CCy2 Fx rate to HorizonDate
' VolToExpiry: The volatility of the CCy1/CCy2 Fx rate to ExpiryDate
' DF1ToHrzn : The CCY1 discount factor to HorizonDate
' DF1ToExpiry: The CCY1 discount factor to Expiry date
' DF2ToHrzn : The CCY2 discount factor to HorizonDate
' DF2ToExpiry: The CCY2 discount factor to Expiry date
' CCYList   : A one-column list of all currencies in the portfolio, excluding the BaseCCy
' Forwards  : Corresponding forward rates CCY/BaseCCY to the HorizonDate
' Vols      : The volatilities of the CCy/BaseCCy rates to the HorizonDate
' Correlations: The correlation matrix for the CCy/BaseCCY rates to the HorizonDate
' ControlString: A comma-delimited string to determine the function's return. Defaults, if omitted to
'             "PFE". Recognised substrings are:
'             PFE,PathValues,AveragePathValue,DF1ToHorizon,DF1ToExpiry,DF2ToHorizon,DF2ToExpiry,ForwardVol,Forwards,Vols,Correlations,MaxControlString
' -----------------------------------------------------------------------------------------------------------------------
Function sFxOptionPFE(AnchorDate As Long, HorizonDate As Double, BaseCCY As String, NumPaths As Long, RNGName As String, _
        Confidence As Variant, IsShortfall As Boolean, BuySell As Variant, OptionStyle As Variant, Ccy1 As Variant, Ccy2 As Variant, _
        Amount1 As Variant, Amount2 As Variant, Expiry As Variant, Strike As Variant, VolToHrzn As Variant, _
        VolToExpiry As Variant, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn, DF2ToExpiry, CCyList, _
        Forwards, Vols, Correlations, Optional ControlString As String = "PFE")
Attribute sFxOptionPFE.VB_Description = "Calculates the PFE on a portfolio of FxOptions"
Attribute sFxOptionPFE.VB_ProcData.VB_Invoke_Func = " \n29"

1         On Error GoTo ErrHandler

          Dim AccurateStrike
          Dim Oss() As EnmOptStyle

2         Force2DArrayRMulti BuySell, OptionStyle, Ccy1, Ccy2, Amount1, Amount2, Expiry, Strike, VolToHrzn, _
              VolToExpiry, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn, DF2ToExpiry, CCyList, Forwards, Vols, Correlations, Confidence

3         Oss = StringsToOptStyle(OptionStyle, True)

4         fxopCheckInputs AnchorDate, HorizonDate, BaseCCY, NumPaths, _
              Confidence, BuySell, Oss, Ccy1, Ccy2, _
              Amount1, Amount2, Expiry, Strike, AccurateStrike, VolToHrzn, _
              VolToExpiry, DF1ToHrzn, DF1ToExpiry, _
              DF2ToHrzn, DF2ToExpiry, CCyList, _
              Forwards, Vols, Correlations, 2, False

5         sFxOptionPFE = sFxOptionPFECore(AnchorDate, HorizonDate, BaseCCY, NumPaths, RNGName, _
              Confidence, IsShortfall, BuySell, Oss, Ccy1, Ccy2, Amount1, Amount2, _
              Expiry, AccurateStrike, VolToHrzn, VolToExpiry, DF1ToHrzn, DF1ToExpiry, _
              DF2ToHrzn, DF2ToExpiry, CCyList, Forwards, Vols, Correlations)

6         Exit Function
ErrHandler:
7         sFxOptionPFE = "#sFxOptionPFE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : fxopCheckInputs
' Author    : Philip Swannell
' Date      : 04-Jul-2015
' Purpose   : Sanity checking of input variables for fxOptionPFE, also sets ByRef arguments NT and NCCYs
'       Mode = 2: Check all arguments
'       Mode = 1: Check only arguments AnchorDate to Strike
'       Mode = 0: Check arguments BuySell to Strike only (i.e arguments that define the trades)
' RelaxedForForwards if True then for Forwards we allow: a) CCY1 = CCY2, b) arbitrary sign for Amount1 and Amount2
' -----------------------------------------------------------------------------------------------------------------------
Public Sub fxopCheckInputs(AnchorDate As Long, HorizonDate As Double, BaseCCY As String, NumPaths As Long, _
        Confidence As Variant, BuySell As Variant, Oss() As EnmOptStyle, Ccy1 As Variant, Ccy2 As Variant, _
        Amount1 As Variant, Amount2 As Variant, Expiry As Variant, Strike As Variant, _
        ByRef AccurateStrike As Variant, VolToHrzn As Variant, _
        VolToExpiry As Variant, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn, DF2ToExpiry, CCyList, _
        Forwards, Vols, Correlations, Mode As Long, RelaxedForForwards As Boolean)

          ' Note this function is called from the Cayley workbook, so we can't make in Private.

          Const ErrBuySell = "BuySell must be a column array with elements ""B"" or ""S"""
          Const ErrOptionStyle = "OptionStyle must be a column array with allowed elements Call (or C), Put (or P), Buy (or B), Sell (or S), Up Digital (or UD), Down Digital (or DD)"""
          Const ErrNumMCPaths = "NumPaths must be between 1 and 100,000"
          Const MAXPaths = 1000000
          Dim ErrConfidence As String
          Const ErrNumRows = "All of the arguments BuySell, OptionStyle, CCY1, CCY2, Amount1, Amount2, Expiry, Strike, VolToHrzn, VolToExpiry, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn and DF2ToExpiry must have the same number of rows"
          Const ErrCcyList = "Invalid CCyList. Elements must be strings, all must be unique and different from BaseCcy"
          Const ErrCCY1 = "CCy1 must be a column array of currency codes. Each element must either be equal to BaseCcy or be a currency code that appears in CcyList"
          Const ErrCCY2 = "CCy2 must be a column array of currency codes. Each element must either be equal to BaseCcy or be a currency code that appears in CcyList"
          Const ErrAmount1 = "Amount1 must be a column array of numbers giving the amounts of currency1"
          Const ErrAmount2 = "Amount2 must be a column array of numbers giving the amounts of currency1"
          Const ErrStrike = "Strike must be a column array of numbers giving the strikes of the options"
          Const StrikeEpsilon = 0.0001
          Const ErrExpiry = "Expiry must be a column array of dates giving the expiry dates of the options"
          Const ErrVolToHrzn = "VolToHrzn, must be a column array of numbers (1 element for each trade) giving the volatility of each currency pair to the HorizonDate"
          Const ErrVolToExpiry = "VolToExpiry must be a column array of numbers (1 element for each trade) giving the volatility of each currency pair to the HorizonDate"
          Const ErrDF1ToExpiry = "ErrDF1ToExpiry must be a column array of positive numbers giving the CCY1 discount factor to the expiry date of each option"
          Const ErrDF1ToHrzn = "ErrDF1ToHrzn must be a column array of positive numbers (1 element for each trade) giving the CCY1 discount factor to the HorizonDate"
          Const ErrDF2ToExpiry = "ErrDF2ToExpiry must be a column array of positive numbers giving the CCY2 discount factor to the expiry date of each option"
          Const ErrDF2ToHrzn = "ErrDF2ToHrzn must be a column array of positive numbers (1 element for each trade) giving the CCY2 discount factor to the HorizonDate"
          Const ErrForwards = "Forwards must be a column array of positive numbers with the same number of elements of CCYList. Each element gives the forward Fx rate CCY/BaseCCY to the HorizonDate."
          Const ErrVols = "Vols must be a column array of positive numbers with the same number of elements of CCYList. Each element gives the volatility of the Fx rate CCY/BaseCCY to the HorizonDate."
          Const ErrCorrelations = "Correlations must be a symmetric positive definite matrix giving the correlations of the CCyList/BaseCcy Fx rates to the HorizonDate"
          Dim i As Long
          Dim j As Long
          Dim NCcys As Long
          Dim NT As Long
          Dim RelaxedForThisTrade As Boolean
          Dim v As Variant

1         On Error GoTo ErrHandler

          'Start input checking, may want to have mode when input checking is omited
2         NT = sNRows(BuySell)
3         If Mode = 2 Then NCcys = sNRows(CCyList)
4         If Mode > 0 Then
5             If HorizonDate < AnchorDate Then Throw "HorizonDate must be after AnchorDate"
6             If NumPaths < 1 Or NumPaths > MAXPaths Then Throw ErrNumMCPaths
7             If sNCols(Confidence) <> 1 Then Throw ErrConfidence
8             Force2DArray Confidence
9             For Each v In Confidence
10                If Not IsNumber(v) Then Throw ErrConfidence
11                If v < 1 / NumPaths Or v > 1 - 1 / NumPaths Then
12                    ErrConfidence = "ConfidenceLevels must be a column array of values between 1/" + CStr(NumPaths) + " and " + CStr(NumPaths - 1) + "/" + CStr(NumPaths)
13                    Throw ErrConfidence
14                End If
15            Next
16        End If
17        If sNCols(BuySell) <> 1 Then Throw ErrBuySell
18        For Each v In BuySell
19            If UCase$(CStr(v)) <> "B" And UCase$(CStr(v)) <> "S" Then Throw ErrBuySell
20        Next
21        If sNCols(Oss) <> 1 Then Throw ErrOptionStyle
22        If sNRows(Oss) <> NT Then Throw ErrNumRows
23        For i = 1 To NT
24            If Oss(i, 1) = OptStyleSell Then Throw "Unexpected error: OptStyleSell is not allowed"
25        Next i

26        If Mode = 2 Then
27            If sNCols(CCyList) <> 1 Then Throw ErrCcyList
28            For Each v In CCyList
29                If VarType(v) <> vbString Then Throw ErrCcyList
30            Next v

31            If sNRows(sRemoveDuplicates(CCyList)) <> sNRows(CCyList) Then Throw ErrCcyList
32            If IsNumber(sMatch(BaseCCY, CCyList)) Then Throw ErrCcyList
33        End If
34        If sNCols(Ccy1) <> 1 Then Throw ErrCCY1
35        If Mode = 2 Then If sNRows(sCompareTwoArrays(Ccy1, sArrayStack(BaseCCY, CCyList), "In1AndNotIn2")) > 1 Then Throw ErrCCY1
36        If sNRows(Ccy1) <> NT Then Throw ErrNumRows
37        If sNCols(Ccy2) <> 1 Then Throw ErrCCY2
38        If Mode = 2 Then If sNRows(sCompareTwoArrays(Ccy2, sArrayStack(BaseCCY, CCyList), "In1AndNotIn2")) > 1 Then Throw ErrCCY2
39        If sNRows(Ccy2) <> NT Then Throw ErrNumRows
40        For Each v In Amount1
41            If Not IsNumberOrDate(v) Then Throw ErrAmount1
42        Next v
43        If sNRows(Amount1) <> NT Then Throw ErrNumRows
44        If sNCols(Amount1) <> 1 Then Throw ErrAmount1
45        For Each v In Amount2
46            If Not IsNumberOrDate(v) Then Throw ErrAmount2
47        Next v
48        If sNRows(Amount2) <> NT Then Throw ErrNumRows
49        If sNCols(Amount2) <> 1 Then Throw ErrAmount2
50        For Each v In Expiry
51            If Not IsNumberOrDate(v) Then Throw ErrExpiry
              'Should we force expiry tyo be greater than Horizon date or just silently exclude when not?
52        Next v
53        If sNRows(Expiry) <> NT Then Throw ErrNumRows
54        If sNCols(Expiry) <> 1 Then Throw ErrNumRows
55        For Each v In Strike
56            If Not IsNumberOrDate(v) Then Throw ErrStrike
57        Next v
58        If sNRows(Strike) <> NT Then Throw ErrNumRows
59        If sNCols(Strike) <> 1 Then Throw ErrNumRows

60        AccurateStrike = sReshape(0, NT, 1)

61        For i = 1 To NT
62            If VarType(Ccy1(i, 1)) <> vbString Then Throw "CCY1 must be a column array of strings. Non string detected at element " + CStr(i)
63            If Len(Ccy1(i, 1)) <> 3 Then Throw "Currency string must be of length 3, but in CCY1 the string at position " + CStr(i) + " is '" + Ccy1(i, 1) + "' i.e. of length " + CStr(Len(Ccy1(i, 1)))
64            If VarType(Ccy2(i, 1)) <> vbString Then Throw "CCY2 must be a column array of strings. Non string detected at element " + CStr(i)
65            If Len(Ccy2(i, 1)) <> 3 Then Throw "Currency string must be of length 3, but in CCY2 the string at position " + CStr(i) + " is '" + Ccy2(i, 1) + "' i.e. of length " + CStr(Len(Ccy2(i, 1)))

66            RelaxedForThisTrade = RelaxedForForwards And (Oss(i, 1) = OptStyleBuy Or Oss(i, 1) = OptStyleSell)
67            If Not RelaxedForThisTrade Then
68                If UCase$(Ccy1(i, 1)) = UCase$(Ccy2(i, 1)) Then
69                    Throw "For options, CCY1  and CCY2 must be different, but at element " + CStr(i) + " they are the same"
70                End If
71                If UCase$(BuySell(i, 1)) = "B" Then
72                    If Amount1(i, 1) < 0 Then Throw "When BuySell is ""B"", Amount1 must be positive or zero, but at element " + CStr(i) + " it is negative"
73                    If Amount2(i, 1) > 0 Then Throw "When BuySell is ""B"", Amount2 must be negative or zero, but at element " + CStr(i) + " it is positive"
74                Else
75                    If Amount1(i, 1) > 0 Then Throw "When BuySell is ""S"", Amount1 must be negative or zero, but at element " + CStr(i) + " it is positive"
76                    If Amount2(i, 1) < 0 Then Throw "When BuySell is ""S"", Amount2 must be positive or zero, but at element " + CStr(i) + " it is negative"
77                End If
78                If Oss(i, 1) <> OptStyleSell Then
79                    If (Amount1(i, 1) = 0 And Amount2(i, 1) <> 0) Or (Amount1(i, 1) <> 0 And Amount2(i, 1) = 0) Then Throw "When Amount1 is zero then Amount2 must be zero and vice versa. This is not the case at element " + CStr(i)
80                End If
81            End If

82            If Amount1(i, 1) <> 0 Then
83                AccurateStrike(i, 1) = -Amount2(i, 1) / Amount1(i, 1)
84                If Not sNearlyEquals(Strike(i, 1), AccurateStrike(i, 1), , StrikeEpsilon) Then Throw "Strike is inconsistent with Amount1 and Amount2 at element " + CStr(i) + ". Strike must be equal to Amount2 divided by -Amount1"
85            End If
86        Next i

87        If Mode = 2 Then

88            For Each v In VolToHrzn
89                If Not IsNumber(v) Then Throw ErrVolToHrzn
90                If v < 0 Then Throw ErrVolToHrzn
91            Next v
92            If sNRows(VolToHrzn) <> NT Then Throw ErrNumRows
93            If sNCols(VolToHrzn) <> 1 Then Throw ErrVolToHrzn

94            For Each v In VolToExpiry
95                If Not IsNumber(v) Then Throw ErrVolToExpiry
96                If v < 0 Then Throw ErrVolToExpiry
97            Next v
98            If sNRows(VolToExpiry) <> NT Then Throw ErrNumRows
99            If sNCols(VolToExpiry) <> 1 Then Throw ErrVolToExpiry

100           For Each v In DF1ToHrzn
101               If Not IsNumber(v) Then Throw ErrDF1ToHrzn
102               If v < 0 Then Throw ErrDF1ToHrzn
103           Next v
104           If sNRows(DF1ToHrzn) <> NT Then Throw ErrNumRows
105           If sNCols(DF1ToHrzn) <> 1 Then Throw ErrDF1ToHrzn

106           For Each v In DF1ToExpiry
107               If Not IsNumber(v) Then Throw ErrDF1ToExpiry
108               If v < 0 Then Throw ErrDF1ToExpiry
109           Next v
110           If sNRows(DF1ToExpiry) <> NT Then Throw ErrNumRows
111           If sNCols(DF1ToExpiry) <> 1 Then Throw ErrDF1ToExpiry

112           For Each v In DF2ToHrzn
113               If Not IsNumber(v) Then Throw ErrDF2ToHrzn
114               If v < 0 Then Throw ErrDF2ToHrzn
115           Next v
116           If sNRows(DF2ToHrzn) <> NT Then Throw ErrNumRows
117           If sNCols(DF2ToHrzn) <> 1 Then Throw ErrDF2ToHrzn

118           For Each v In DF2ToExpiry
119               If Not IsNumber(v) Then Throw ErrDF2ToExpiry
120               If v < 0 Then Throw ErrDF2ToExpiry
121           Next v
122           If sNRows(DF2ToExpiry) <> NT Then Throw ErrNumRows
123           If sNCols(DF2ToExpiry) <> 1 Then Throw ErrDF2ToExpiry

124           For Each v In Forwards
125               If Not IsNumber(v) Then Throw ErrForwards
126               If v < 0 Then Throw ErrForwards
127           Next v
128           If sNRows(Forwards) <> NCcys Then Throw ErrForwards
129           If sNCols(Forwards) <> 1 Then Throw ErrForwards

130           For Each v In Vols
131               If Not IsNumber(v) Then Throw ErrVols
132               If v < 0 Then Throw ErrVols
133           Next v
134           If sNRows(Vols) <> NCcys Then Throw ErrVols
135           If sNCols(Vols) <> 1 Then Throw ErrVols

136           If sNRows(Correlations) <> NCcys Then Throw ErrCorrelations
137           If sNCols(Correlations) <> NCcys Then Throw ErrCorrelations
138           For Each v In Correlations
139               If Not IsNumber(v) Then Throw ErrCorrelations
140               If v < 0 Or v > 1 Then Throw ErrCorrelations
141           Next v

142           For i = 1 To NCcys
143               If Correlations(i, i) <> 1 Then Throw ErrCorrelations
144           Next i

145           For i = 1 To NCcys
146               For j = 1 To i - 1
147                   If Correlations(i, j) <> Correlations(j, i) Then Throw ErrCorrelations
148               Next j
149           Next i

150       End If

151       Exit Sub
ErrHandler:
152       Throw "#fxopCheckInputs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFxOptionPFE2
' Author    : Philip Swannell
' Date      : 06-Jul-2015
' Purpose   : More convenient version of sFxOption - market data is passed as the name of
'             a standard-Format market workbook
' -----------------------------------------------------------------------------------------------------------------------
Function sFxOptionPFE2(HorizonDate As Double, BaseCCY As String, NumPaths As Long, RNGName As String, MarketBookName As String, _
        WithShocks As Boolean, Confidence As Variant, IsShortfall As Boolean, BuySell As Variant, OptionStyle As Variant, Ccy1 As Variant, Ccy2 As Variant, _
        Amount1 As Variant, Amount2 As Variant, Expiry As Variant, Strike As Variant, Optional ControlString As String = "PFE")
Attribute sFxOptionPFE2.VB_Description = "Calculates the PFE on a portfolio of FxOptions. For convenience, market data passed as the name of a workbook containing discount factors, Fx rates, volatilities and correlations."
Attribute sFxOptionPFE2.VB_ProcData.VB_Invoke_Func = " \n29"

          Dim AccurateStrike
          Dim AnchorDate As Long
          Dim CCyList
          Dim Correlations
          Dim DF1ToExpiry
          Dim DF1ToHrzn
          Dim DF2ToExpiry
          Dim DF2ToHrzn
          Dim Forwards
          Dim i As Long
          Dim Mode As Long
          Dim NumCcys As Long
          Dim NumOpts As Long
          Dim OS As EnmOptStyle
          Dim Oss() As EnmOptStyle
          Dim Temp
          Dim Vols
          Dim VolToExpiry As Variant
          Dim VolToHrzn As Variant

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti BuySell, OptionStyle, Ccy1, Ccy2, Amount1, Amount2, Expiry, Strike, Confidence
3         Oss = StringsToOptStyle(OptionStyle, True)
4         Mode = 1
5         AnchorDate = ThrowIfError(sMarketAnchorDate(MarketBookName))
6         fxopCheckInputs AnchorDate, HorizonDate, BaseCCY, NumPaths, Confidence, BuySell, Oss, Ccy1, Ccy2, Amount1, Amount2, Expiry, Strike, AccurateStrike, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Mode, False

7         NumOpts = sNRows(BuySell)
8         Temp = sReshape(0, NumOpts, 1)
9         VolToHrzn = Temp: VolToExpiry = Temp: DF1ToHrzn = Temp: DF1ToExpiry = Temp: DF2ToHrzn = Temp: DF2ToExpiry = Temp

10        For i = 1 To NumOpts
11            OS = Oss(i, 1)
12            If OS <> OptStyleBuy Then        'OptStyleSell is never encountered
13                VolToHrzn(i, 1) = FirstElement(sMarketFxVol(CStr(Ccy1(i, 1)), CStr(Ccy2(i, 1)), HorizonDate, MarketBookName, False, WithShocks))
14                VolToExpiry(i, 1) = FirstElement(sMarketFxVol(CStr(Ccy1(i, 1)), CStr(Ccy2(i, 1)), Expiry(i, 1), MarketBookName, False, WithShocks))
15            End If
16            DF1ToHrzn(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy1(i, 1)), HorizonDate, MarketBookName))
17            DF1ToExpiry(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy1(i, 1)), Expiry(i, 1), MarketBookName))
18            DF2ToHrzn(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy2(i, 1)), HorizonDate, MarketBookName))
19            DF2ToExpiry(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy2(i, 1)), Expiry(i, 1), MarketBookName))
20        Next i

21        CCyList = sCompareTwoArrays(sArrayStack(Ccy1, Ccy2), BaseCCY, "In1AndNotIn2,NoHeaders")
22        NumCcys = sNRows(CCyList)
23        Temp = sReshape(0, NumCcys, 1)
24        Forwards = Temp: Vols = Temp
25        For i = 1 To NumCcys
26            Forwards(i, 1) = FirstElement(sMarketFxForwardRates(HorizonDate, CStr(CCyList(i, 1)), BaseCCY, MarketBookName, WithShocks))
27            Vols(i, 1) = FirstElement(sMarketFxVol(CStr(CCyList(i, 1)), BaseCCY, HorizonDate, MarketBookName, , WithShocks))
28        Next i
29        Correlations = ThrowIfError(sMarketCorrelationMatrix(CCyList, BaseCCY, MarketBookName))

30        sFxOptionPFE2 = sFxOptionPFECore(AnchorDate, HorizonDate, BaseCCY, NumPaths, RNGName, Confidence, IsShortfall, BuySell, Oss, Ccy1, Ccy2, Amount1, Amount2, Expiry, AccurateStrike, VolToHrzn, VolToExpiry, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn, DF2ToExpiry, CCyList, Forwards, Vols, Correlations, , ControlString)

31        Exit Function
ErrHandler:
32        sFxOptionPFE2 = "#sFxOptionPFE2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FirstElement
' Author    : Philip Swannell
' Date      : 06-Jul-2015
' Purpose   : Assumes Data is either a 2-d array or not an array
' -----------------------------------------------------------------------------------------------------------------------
Private Function FirstElement(Data)
1         If VarType(Data) < vbArray Then
2             If VarType(Data) = vbString Then If Left$(Data, 1) = "#" Then If Right$(Data, 1) = "!" Then Throw CStr(Data)
3             FirstElement = Data
4         Else
5             FirstElement = Data(1, 1)
6         End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFxOptionPFECore
' Author    : Philip Swannell
' Date      : 07-Jul-2015
' Purpose   : Code which does the "work", no error checking of inputs so should not be called directly by users
'  The function returns the PFE of the sum of a) The trades passed in via arguments BuySell thru Strike, and
'                                             b) Cashflows on the horizon date of BaseCCyHolding and FgnCcyHoldings (currencies correspond to those in CCYList
'  b) allows for a large number of FxForwards to be efficiently represented - possible because discounting is non stochastic.
' -----------------------------------------------------------------------------------------------------------------------
Function sFxOptionPFECore(AnchorDate As Long, HorizonDate As Double, BaseCCY As String, NumPaths As Long, RNGName As String, _
        Confidence As Variant, IsShortfall As Boolean, BuySell As Variant, Oss() As EnmOptStyle, Ccy1 As Variant, Ccy2 As Variant, _
        Amount1 As Variant, Amount2 As Variant, Expiry As Variant, Strike As Variant, VolToHrzn As Variant, _
        VolToExpiry As Variant, DF1ToHrzn, DF1ToExpiry, DF2ToHrzn, DF2ToExpiry, CCyList, _
        Forwards, Vols, Correlations, Optional ByVal CorrelatedNormals, Optional ControlString As String = "PFE", _
        Optional BaseCCYHolding As Double, Optional FgnCCYHoldings As Variant)

1         On Error GoTo ErrHandler

          Dim NT As Long        'Number of Trades,
          Dim i As Long
          Dim j As Long
          Dim NCcys As Long
          Dim OptTime() As Double        'NT rows, Time argument for each call to BlackScholes
          Dim ForwardVol() As Double        ' NT rows, Vol argument for each call to BlackScholes
          Dim ForwardRatio() As Double        'NT rows. For each trade gives the ratio ForwardRateToExpiry/ForwardRateToHrzn
          Dim CCy1Index As Variant        'NT rows. For each trade gives the index of CCY1 into CcyList or 0 if CCY1 is BaseCCY
          Dim CCy2Index As Variant        'NT rows. For each trade gives the index of CCY2 into CcyList or 0 if CCY2 is BaseCCY
          Dim Cholesky As Variant
          Dim CorrelatedLogNormals As Variant
          Dim CurrentMeans As Variant
          Dim CurrentStDev As Variant
          Dim DesiredMeans As Variant
          Dim DesiredStDev As Variant
          Dim Normals As Variant
          Dim PathValues() As Double
          Dim Rebase As Boolean
          Dim RebasedCorrelatedNormals As Variant
          Dim TAtoH As Double
          Dim Weights() As Double

2         If IsEmpty(BuySell) Then
3             NT = 0
4         Else
5             NT = sNRows(BuySell)
6         End If
7         NCcys = sNRows(CCyList)

8         If NT > 0 Then
9             ReDim OptTime(1 To NT, 1 To 1)
10            For i = 1 To NT
11                OptTime(i, 1) = (Expiry(i, 1) - HorizonDate) / 365
12            Next i

13            ReDim ForwardVol(1 To NT, 1 To 1)
14            For i = 1 To NT
15                If Oss(i, 1) <> OptStyleBuy Then
16                    If Expiry(i, 1) <= HorizonDate Then
17                        ForwardVol(i, 1) = 0
18                    Else
19                        ForwardVol(i, 1) = (VolToExpiry(i, 1) ^ 2 * (Expiry(i, 1) - AnchorDate) - VolToHrzn(i, 1) ^ 2 * (HorizonDate - AnchorDate)) / (Expiry(i, 1) - HorizonDate)
20                        If ForwardVol(i, 1) < 0 Then
21                            Throw "Trade " + CStr(i) + " has negative forward variance for " + Ccy1(i, 1) + Ccy2(i, 1) + _
                                  " from Horizon to Expiry (" + Format$(HorizonDate, "dd-mmm-yyyy") + ", " + Format$(VolToHrzn(i, 1), "0.00%") + " to " + _
                                  Format$(Expiry(i, 1), "dd-mmm-yyyy") + ", " + Format$(VolToExpiry(i, 1), "0.00%") + ")"
22                        End If
23                        ForwardVol(i, 1) = ForwardVol(i, 1) ^ 0.5
24                    End If
25                End If
26            Next i

27            ReDim ForwardRatio(1 To NT, 1 To 1)
28            For i = 1 To NT
29                ForwardRatio(i, 1) = (DF1ToExpiry(i, 1) / DF1ToHrzn(i, 1)) / (DF2ToExpiry(i, 1) / DF2ToHrzn(i, 1))
30            Next i
31            ReDim Weights(1 To NT, 1 To 1)
32            For i = 1 To NT
33                Weights(i, 1) = Amount1(i, 1) * DF2ToExpiry(i, 1) / DF2ToHrzn(i, 1)
34            Next i

35            CCy1Index = sArraySubtract(sMatch(Ccy1, sArrayStack(BaseCCY, CCyList)), 1)
36            CCy2Index = sArraySubtract(sMatch(Ccy2, sArrayStack(BaseCCY, CCyList)), 1)
37            If NT = 1 Then
38                Force2DArrayRMulti CCy1Index, CCy2Index
39            End If

40        End If

41        If IsEmpty(CorrelatedNormals) Or IsMissing(CorrelatedNormals) Then
42            Normals = ThrowIfError(sRandomVariable(NumPaths, NCcys, "Normal", RNGName))
43            If sArraysIdentical(Correlations, sIdentityMatrix(sNRows(Forwards))) Then
44                CorrelatedNormals = Normals
45            Else
46                Cholesky = ThrowIfError(sCholesky(Correlations))
47                CorrelatedNormals = Application.WorksheetFunction.MMult(Normals, sArrayTranspose(Cholesky))
48            End If
49        Else
50            If sNCols(CorrelatedNormals) <> NCcys Then Throw "If provided, CorrelatedNormals must have the same number of columns as there are elements in CCyList"
51            If sNRows(CorrelatedNormals) <> NumPaths Then Throw "If provided, the number of rows of CorrelatedNormals must be the number given by NumPaths"
52        End If

53        Rebase = InStr(LCase$(RNGName), "sobol") = 0        'We don't correct mean or standard dev for sobol sequences
54        TAtoH = (HorizonDate - AnchorDate) / 365

55        DesiredMeans = sArrayTranspose(sArraySubtract(sArrayLog(Forwards), sArrayMultiply(0.5, Vols, Vols, TAtoH)))
56        DesiredStDev = sArrayTranspose(sArrayMultiply(Vols, sArrayPower(TAtoH, 0.5)))
57        If Rebase Then
58            CurrentMeans = sColumnMean(CorrelatedNormals)
59            CurrentStDev = sColumnStDev(CorrelatedNormals)
60            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(sArraySubtract(CorrelatedNormals, CurrentMeans), sArrayDivide(DesiredStDev, CurrentStDev)), DesiredMeans)
61        Else
62            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(CorrelatedNormals, DesiredStDev), DesiredMeans)
63        End If

64        CorrelatedLogNormals = sArrayExp(RebasedCorrelatedNormals)

65        ReDim PathValues(1 To NumPaths, 1 To 1)

66        If NT > 0 Then
              Dim CCY1ToBase As Double
              Dim CCY2ToBase As Double
              Dim Fwd As Double
              Dim Indx1 As Long
              Dim Indx2 As Long
              Dim OS As EnmOptStyle
              Dim Strk As Double
              Dim t As Double
              Dim Vol As Double
67            For i = 1 To NumPaths
68                For j = 1 To NT
69                    t = OptTime(j, 1)
70                    If t >= 0 Then        'Expired options treated as having zero value after their expiry date, but intrinsic on their expiry date.
71                        OS = Oss(j, 1)
72                        Indx1 = CCy1Index(j, 1)
73                        Indx2 = CCy2Index(j, 1)
74                        If Indx1 = 0 Then
75                            CCY1ToBase = 1
76                        Else
77                            CCY1ToBase = CorrelatedLogNormals(i, Indx1)
78                        End If
79                        If Indx2 = 0 Then
80                            CCY2ToBase = 1
81                        Else
82                            CCY2ToBase = CorrelatedLogNormals(i, Indx2)
83                        End If
84                        Fwd = CCY1ToBase / CCY2ToBase * ForwardRatio(j, 1)        ' Rates are non-stochastic, so the forward to the expiry date _
                                                                                      is just the forward to the horizon date multiplied by a ratio _
                                                                                      which for a given trade is constant across all paths.
85                        Strk = Strike(j, 1)
86                        Vol = ForwardVol(j, 1)
87                        If OS = OptStyleBuy Then
                              'For forwards, we use Amount1 and Amount2 whereas for Options we look at Amount1 and Strike
88                            PathValues(i, 1) = PathValues(i, 1) + Amount1(j, 1) * DF1ToExpiry(j, 1) / DF1ToHrzn(j, 1) * CCY1ToBase + _
                                  Amount2(j, 1) * DF2ToExpiry(j, 1) / DF2ToHrzn(j, 1) * CCY2ToBase
89                        Else
90                            If Weights(j, 1) <> 0 Then
91                                PathValues(i, 1) = PathValues(i, 1) + bsCore(OS, Fwd, Strk, Vol, t) * Weights(j, 1) * CCY2ToBase        ' The option pays out in Ccy2, so we need to translate to the Base (reporting) currency.
92                            End If
93                        End If

94                    End If
95                Next j
96            Next i
97        End If

98        If Not (IsEmpty(FgnCCYHoldings) Or IsMissing(FgnCCYHoldings)) Then
              Dim MmultRes As Variant
99            MmultRes = Application.WorksheetFunction.MMult(CorrelatedLogNormals, FgnCCYHoldings)
100           For i = 1 To NumPaths
101               PathValues(i, 1) = PathValues(i, 1) + BaseCCYHolding + MmultRes(i, 1)
102           Next i
103       End If

          Const MaxControlString = "PFE,PathValues,AveragePathValue,DF1ToHorizon,DF1ToExpiry,DF2ToHorizon,DF2ToExpiry,ForwardVol,Forwards,Vols,Correlations,MaxControlString"
          Dim PFE
104       If InStr(ControlString, "PFE") > 0 Then
              Dim NumCL As Long
105           NumCL = sNRows(Confidence)
106           PFE = sReshape(0, NumCL, 1)
107           For i = 1 To NumCL
108               If IsShortfall Then
                      Dim CutOff As Double
109                   CutOff = Application.WorksheetFunction.Percentile_Exc(PathValues, CDbl(Confidence(i, 1)))
110                   PFE(i, 1) = AverageIf(PathValues, CutOff)

111               Else
112                   PFE(i, 1) = Application.WorksheetFunction.Percentile_Exc(PathValues, CDbl(Confidence(i, 1)))
113               End If
114           Next i
115       End If

116       If ControlString = "PFE" Then
117           sFxOptionPFECore = PFE
118       Else
              Dim ControlStringArray
              Dim Result
119           ControlStringArray = sTokeniseString(ControlString)
120           Force2DArray ControlStringArray
121           Result = CreateMissing()
122           For i = 1 To sNRows(ControlStringArray)
123               Select Case UCase$(ControlStringArray(i, 1))
                      Case "PFE"
124                       Result = sArrayRange(Result, PFE)
125                   Case UCase$("AveragePathValue")
126                       Result = sArrayRange(Result, sColumnSum(PathValues)(1, 1) / NumPaths)
127                   Case UCase$("PathValues")
128                       Result = sArrayRange(Result, PathValues)
129                   Case UCase$("DF1ToHorizon")
130                       Result = sArrayRange(Result, DF1ToHrzn)
131                   Case UCase$("DF1ToExpiry")
132                       Result = sArrayRange(Result, DF1ToExpiry)
133                   Case UCase$("DF2ToHorizon")
134                       Result = sArrayRange(Result, DF2ToHrzn)
135                   Case UCase$("DF2ToExpiry")
136                       Result = sArrayRange(Result, DF2ToExpiry)
137                   Case UCase$("ForwardVol")
138                       Result = sArrayRange(Result, ForwardVol)
139                   Case UCase$("Forwards")
140                       Result = sArrayRange(Result, sArrayRange(sArrayConcatenate(CCyList, BaseCCY), Forwards))
141                   Case UCase$("Vols")
142                       Result = sArrayRange(Result, sArrayRange(sArrayConcatenate(CCyList, BaseCCY), Vols))
143                   Case UCase$("Correlations")
                          Dim CorreltionWithHeaders
                          Dim Headers
144                       Headers = sArrayConcatenate(CCyList, BaseCCY)
145                       CorreltionWithHeaders = sArraySquare(vbNullString, sArrayTranspose(Headers), _
                              Headers, Correlations)
146                       Result = sArrayRange(Result, CorreltionWithHeaders)
147                   Case UCase$("MaxControlString")
148                       Result = sArrayRange(Result, MaxControlString)
149               End Select
150           Next i
151           If IsMissing(Result) Then
152               Result = "#Nothing in ControlString recognised maximum ControlString is " + MaxControlString + "!"
153           End If
154           sFxOptionPFECore = Result
155       End If

156       Exit Function
ErrHandler:
157       Throw "#sFxOptionPFECore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AverageIf
' Author    : Philip Swannell
' Date      : 07-Nov-2016
' Purpose   : Sub-routine of sFxOptionPFECore - can't use
'             Application.WorksheetFunction.AverageIf since that requires first agument as Range :-(
' -----------------------------------------------------------------------------------------------------------------------
Private Function AverageIf(Values, CutOff)
          Dim i As Long
          Dim N As Long
          Dim Total As Double

1         On Error GoTo ErrHandler
2         For i = 1 To sNRows(Values)
3             If Values(i, 1) >= CutOff Then
4                 N = N + 1
5                 Total = Total + Values(i, 1)
6             End If
7         Next i

8         If N = 0 Then Throw "CutOff too high. None of " + CStr(sNRows(Values)) & " path values exceed or match CutOff of " + CStr(CutOff)
9         AverageIf = Total / N
10        Exit Function
ErrHandler:
11        Throw "#AverageIf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFxOptionValues
' Author    : Philip Swannell
' Date      : 08-Jul-2015
' Purpose   : Takes trade data in same Format as sFxOptionPFE2 and returns a column array of trade values
' -----------------------------------------------------------------------------------------------------------------------
Function sFxOptionValues(BaseCCY As String, MarketBookName As String, _
        WithShocks As Boolean, BuySell As Variant, OptionStyle As Variant, Ccy1 As Variant, Ccy2 As Variant, _
        Amount1 As Variant, Amount2 As Variant, Expiry As Variant, Strike As Variant)

          Dim AccurateStrike As Variant
          Dim AnchorDate As Long
          Dim CCY1ToBase
          Dim CCY2ToBase
          Dim DF1ToExpiry
          Dim DF2ToExpiry
          Dim Forwards
          Dim i As Long
          Dim Mode As Long
          Dim NumOpts As Long
          Dim OS As EnmOptStyle
          Dim Oss() As EnmOptStyle
          Dim Temp
          Dim Times
          Dim ValuesInBaseCCY
          Dim VolToExpiry As Variant

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti BuySell, OptionStyle, Ccy1, Ccy2, Amount1, Amount2, Expiry, Strike
3         AnchorDate = ThrowIfError(sMarketAnchorDate(MarketBookName))
4         NumOpts = sNRows(BuySell)

          'We secretly allow OptionStyle to be passed not as strings but as EnmOptStyle, so we don't call StringsToOptStyles, which _
           insists its argument is an array of strings
5         ReDim Oss(1 To NumOpts, 1 To 1)
6         For i = 1 To NumOpts
7             If IsNumber(OptionStyle(i, 1)) Then
8                 Oss(i, 1) = OptionStyle(i, 1)
9             Else
10                Oss(i, 1) = StringToOptStyle(OptionStyle(i, 1), True)
11            End If
12        Next i

13        Mode = 1        ' controls what work fxopCheckInputs does
14        fxopCheckInputs AnchorDate, AnchorDate + 1, BaseCCY, 100, 0.95, BuySell, Oss, Ccy1, Ccy2, Amount1, Amount2, Expiry, Strike, AccurateStrike, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Mode, False

15        Temp = sReshape(0, NumOpts, 1)
16        VolToExpiry = Temp: DF1ToExpiry = Temp: DF2ToExpiry = Temp: ValuesInBaseCCY = Temp

17        Forwards = ThrowIfError(sMarketFxForwardRates2(Expiry, Ccy1, Ccy2, MarketBookName, WithShocks))
18        CCY1ToBase = sMarketFxPerBaseCcy(Ccy1, BaseCCY, MarketBookName, WithShocks)
19        CCY2ToBase = sMarketFxPerBaseCcy(Ccy2, BaseCCY, MarketBookName, WithShocks)
20        Times = sArrayDivide(sArraySubtract(Expiry, AnchorDate), 365)

21        For i = 1 To NumOpts
22            If Times(i, 1) < 0 Then
23                ValuesInBaseCCY(i, 1) = 0
24            Else
25                OS = Oss(i, 1)
26                DF1ToExpiry(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy1(i, 1)), Expiry(i, 1), MarketBookName))
27                DF2ToExpiry(i, 1) = FirstElement(sMarketDiscountFactor(CStr(Ccy2(i, 1)), Expiry(i, 1), MarketBookName))
28                If OS = OptStyleBuy Then
29                    ValuesInBaseCCY(i, 1) = Amount1(i, 1) * DF1ToExpiry(i, 1) * CCY1ToBase(i, 1) + _
                          Amount2(i, 1) * DF2ToExpiry(i, 1) * CCY2ToBase(i, 1)
30                Else

31                    If Amount1(i, 1) = 0 Then
32                        ValuesInBaseCCY(i, 1) = 0
33                    Else
34                        VolToExpiry(i, 1) = FirstElement(sMarketFxVol(CStr(Ccy1(i, 1)), CStr(Ccy2(i, 1)), Expiry(i, 1), MarketBookName, False, WithShocks))
35                        ValuesInBaseCCY(i, 1) = bsCore(OS, Forwards(i, 1), AccurateStrike(i, 1), VolToExpiry(i, 1), Times(i, 1)) * _
                              DF2ToExpiry(i, 1) * Amount1(i, 1) * CCY2ToBase(i, 1)
36                    End If
37                End If
38            End If
39        Next i

40        sFxOptionValues = ValuesInBaseCCY

41        Exit Function
ErrHandler:
42        sFxOptionValues = "#sFxOptionValues (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
