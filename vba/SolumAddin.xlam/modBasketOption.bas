Attribute VB_Name = "modBasketOption"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modBasketOption
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Code to price a European Basket Option by matching the first three moments
'             of the distribution of the basket price to the moments of either a shifted
'             log-normal or a shifted negative log-normal. See for example the paper "American
'             Basket and Spread Option pricing by a Simple Binomial Tree" by S. Borokova et al
'             https://www.feweb.vu.nl/nl/Images/AmericanBasketOptions_tcm96-194676.pdf
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_Target1stMoment As Double
Private m_Target2ndRawMoment As Double
Private m_Target3rdRawMoment As Double
'Private m_Target2ndCentralMoment As Double
'Private m_Target3rdCentralMoment As Double
Private m_BasketSkewIsNegative As Boolean

Private Sub CheckBasketInputs(Forwards, Weights, Vols, Correlations)

1         On Error GoTo ErrHandler
2         Force2DArrayR Forwards: Force2DArrayR Weights: Force2DArrayR Vols: Force2DArrayR Correlations

          Const SizeError = "Forwards, Weights and Vols must be column vectors of the same height"
          Dim c As Variant
          Dim i As Long
          Dim j As Long
          Dim N As Long
3         If sNCols(Forwards) <> 1 Or sNCols(Weights) <> 1 Or sNCols(Vols) <> 1 Then Throw SizeError
4         N = sNRows(Forwards)
5         If sNRows(Weights) <> N Or sNRows(Vols) <> N Then Throw SizeError
6         If sNRows(Correlations) <> N Or sNCols(Correlations) <> N Then Throw "Correlations must be a square matrix with the same number of rows as forwards"
7         For Each c In Forwards
8             If Not IsNumber(c) Then Throw "Forwards must be numbers"
9         Next
10        For Each c In Weights
11            If Not IsNumber(c) Then Throw "Weights must be numbers"
12        Next
13        For Each c In Vols
14            If Not IsNumber(c) Then Throw "Vols must be non-negative numbers"
15            If c < 0 Then Throw "Vols must be non-negative numbers"
16        Next
17        For Each c In Correlations
18            If Not IsNumber(c) Then Throw "Correlations must be numbers"
19            If c < -1 Or c > 1 Then Throw "Correlations must not be outside range -1 to 1"
20        Next
21        For i = 1 To N
22            If Correlations(1, 1) <> 1 Then Throw "Diagonal elements of correlations must be 1"
23        Next i
          'Could check for positive-definite correlation matrix
24        For i = 1 To N
25            For j = 1 To i - 1
26                If Correlations(i, j) <> Correlations(j, i) Then Throw "Correlation matrix must be symmetric"
27            Next j
28        Next i
29        Exit Sub
ErrHandler:
30        Throw "#CheckBasketInputs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketOption
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Prices a European option on a basket of correlated log-normal assets. Pricing is via
'             3-moment-matching to a shifted-log normal asset (rather than multi-variate
'             integration). The function uses sSLNCalibrate and sSLNEuropeanOption.
' Arguments
' Strike    : The Strike of the basket option.
' CP        : The payoff style of the option. Can take values: "C" or "Call", "P" or "Put", "B" or
'             "Buy", "S" or "Sell", "UD" or "Up Digital", "DD" or "Down Digital".
' Forwards  : A column vector of the forward prices of each asset.
' Weights   : A column vector of the weights in the basket of each asset.
' Vols      : A column vector of the log-normal volatilities of the assets. Enter 20% volatility as 0.2.
' Correlations: The correlations between the assets as a matrix. The function checks that the matrix is
'             symmetric, has on-diagonal values of 1 and no values outside the range
'             [-1,1]. It does not check that the matrix is positive definite.
' Time      : The time in years until the basket price is observed.
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketOption(Strike As Double, CP As String, ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double)
Attribute sBasketOption.VB_Description = "Prices a European option on a basket of correlated log-normal assets. Pricing is via 3-moment-matching to a shifted-log normal asset (rather than multi-variate integration). The function uses sSLNCalibrate and sSLNEuropeanOption."
Attribute sBasketOption.VB_ProcData.VB_Invoke_Func = " \n29"
1         On Error GoTo ErrHandler
2         CheckBasketInputs Forwards, Weights, Vols, Correlations        ' also forces Forwards, Weights, Vols and Correlations to be 2-d arrays
          Dim Calibration As Variant

3         Calibration = ThrowIfError(sSLNCalibrate(Forwards, Weights, Vols, Correlations, Time))
4         sBasketOption = sSLNEuropeanOption(Strike, CP, CDbl(Calibration(1, 1)), CDbl(Calibration(2, 1)), CDbl(Calibration(3, 1)), CBool(Calibration(4, 1)))

5         Exit Function
ErrHandler:
6         sBasketOption = "#sBasketOption (line " & CStr(Erl) + "): " & Err.Description & "!"
7         Exit Function
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketOptionMonteCarlo
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Price the Basket Option using a straightforward MonteCarlo - no variance
'             reduction techniques used as we want this method only as check on sBasketOption.
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketOptionMonteCarlo(ByVal Strike, ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double, NumPaths As Long, RNGName As String)
Attribute sBasketOptionMonteCarlo.VB_Description = "Prices a European option on a basket of correlated log-normal assets. Pricing is via Monte Carlo with specified number of paths."
Attribute sBasketOptionMonteCarlo.VB_ProcData.VB_Invoke_Func = " \n29"

1         On Error GoTo ErrHandler
2         CheckBasketInputs Forwards, Weights, Vols, Correlations

3         Force2DArrayR Strike

          Dim BasketPaths
          Dim BuyValue
          Dim CallValue
          Dim Cholesky As Variant
          Dim CorrelatedNormals As Variant
          Dim CurrentMeans As Variant
          Dim CurrentStDev As Variant
          Dim DesiredMeans As Variant
          Dim DesiredStDev As Variant
          Dim DownDigitalValue
          Dim LogNormals
          Dim Normals As Variant
          Dim NumAssets As Long
          Dim PutValue
          Dim RebasedCorrelatedNormals
          Dim Result(1 To 2, 1 To 6)
          Dim SellValue
          Dim UpDigitalValue

4         Cholesky = ThrowIfError(sCholesky(Correlations))
5         NumAssets = sNRows(Forwards)
6         Normals = ThrowIfError(sRandomVariable(NumPaths, NumAssets, "Normal", RNGName))
7         CorrelatedNormals = Application.WorksheetFunction.MMult(Normals, sArrayTranspose(Cholesky))
8         CurrentMeans = sColumnMean(CorrelatedNormals)
9         DesiredMeans = sArrayTranspose(sArraySubtract(sArrayLog(Forwards), sArrayMultiply(0.5, Vols, Vols, Time)))
10        CurrentStDev = sColumnStDev(CorrelatedNormals)
11        DesiredStDev = sArrayTranspose(sArrayMultiply(Vols, sArrayPower(Time, 0.5)))
12        RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(sArraySubtract(CorrelatedNormals, CurrentMeans), sArrayDivide(DesiredStDev, CurrentStDev)), DesiredMeans)
13        LogNormals = sArrayExp(RebasedCorrelatedNormals)
14        BasketPaths = Application.WorksheetFunction.MMult(LogNormals, Weights)
15        CallValue = sColumnMean(sArrayMax(sArraySubtract(BasketPaths, Strike), 0))(1, 1)
16        PutValue = sColumnMean(sArrayMax(sArraySubtract(Strike, BasketPaths), 0))(1, 1)
17        UpDigitalValue = sArrayCount(sArrayGreaterThan(BasketPaths, Strike)) / NumPaths
18        DownDigitalValue = 1 - UpDigitalValue
19        BuyValue = sColumnMean(BasketPaths)(1, 1) - Strike(1, 1)
20        SellValue = -BuyValue
21        Result(1, 1) = "Call": Result(1, 2) = "Put": Result(1, 3) = "Buy": Result(1, 4) = "Sell": Result(1, 5) = "UpDigital": Result(1, 6) = "DownDigital"
22        Result(2, 1) = CallValue: Result(2, 2) = PutValue: Result(2, 3) = BuyValue: Result(2, 4) = SellValue: Result(2, 5) = UpDigitalValue: Result(2, 6) = DownDigitalValue
23        sBasketOptionMonteCarlo = Result

24        Exit Function
ErrHandler:
25        sBasketOptionMonteCarlo = "#sBasketOptionMonteCarlo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNCalibrate
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Calibrates a shifted log-normal distribution to a basket of assets with given
'             forwards, weights, vols and correlation structure. Aim is to match the first
'             three moments of the distribution. First tries Newton-Raphson and if that
'             does not work tries MinPack solver.
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNCalibrate(ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double)
Attribute sSLNCalibrate.VB_Description = "Returns the parameters of a shifted log-normal or shifted negative log-normal distribution so that the distribution's first three moments match those of a weighted basket of correlated log-normal assets."
Attribute sSLNCalibrate.VB_ProcData.VB_Invoke_Func = " \n29"
          Dim Res
1         On Error GoTo ErrHandler

2         CheckBasketInputs Forwards, Weights, Vols, Correlations        ' also forces Forwards, Weights, Vols and Correlations to be 2-d arrays

          'First try using Newton-Raphson
3         Res = sSLNCalibrateNewtonRaphson(Forwards, Weights, Vols, Correlations, Time)
4         If VarType(Res) <> vbString Then
5             sSLNCalibrate = Res
6             Exit Function
7         Else        'If that fails try using Minpack
8             sSLNCalibrate = sSLNCalibrateMinPack(Forwards, Weights, Vols, Correlations, Time)
9         End If

10        Exit Function
ErrHandler:
11        sSLNCalibrate = "#sSLNCalibrate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNCalibrateNewtonRaphson
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Calibrates a shifted log-normal distribution to a basket of assets with given
'             forwards, weights, vols and correlation structure. Uses multi-dimensional Newton Raphson
'             which when it works gives a very accurate solution.
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNCalibrateNewtonRaphson(Forwards, Weights, Vols, Correlations, Time As Double)
1         On Error GoTo ErrHandler
          Dim AdjustmentVector
          Dim DistanceVector
          Dim GuessVec
          Dim i As Long
          Dim inverseJacobian
          Dim Jacobian As Variant
          Dim Moments
          Dim NextGuess
          Const Max_Iterations As Long = 15
          Dim numWeights As Long

          'Only need line below when we are testing this method from a spreadsheet, not when it becomes a Private sub-routine
          '  CheckBasketInputs Forwards, Weights, Vols, Correlations    ' also forces Forwards, Weights, Vols and Correlations to be 2-d arrays

2         numWeights = sNRows(Weights)

3         If sArraysIdentical(Weights, sReshape(0, numWeights, 1)) Then
4             sSLNCalibrateNewtonRaphson = sArrayStack(0, -200, 0, False)
5             Exit Function
6         End If

7         If numWeights = 1 Then
8             sSLNCalibrateNewtonRaphson = sArrayStack(0, _
                  Log(Forwards(1, 1) * Weights(1, 1)) - 0.5 * Vols(1, 1) ^ 2 * Time, _
                  Vols(1, 1) * Sqr(Time), _
                  False)
9             Exit Function
10        End If

11        GuessVec = GuessFirstSolution(Forwards, Weights, Vols, Correlations, Time)
12        ReDim DistanceVector(1 To 1, 1 To 3)

13        ReDim NextGuess(1 To 3)

14        For i = 1 To Max_Iterations
15            Jacobian = sSLNJacobian(CDbl(GuessVec(1)), CDbl(GuessVec(2)), CDbl(GuessVec(3)))

16            inverseJacobian = Application.WorksheetFunction.MInverse(Jacobian)
17            Moments = sSLNRawMoments(CDbl(GuessVec(1)), CDbl(GuessVec(2)), CDbl(GuessVec(3)))
18            DistanceVector(1, 1) = Moments(1, 1) - m_Target1stMoment
19            DistanceVector(1, 2) = Moments(2, 1) - m_Target2ndRawMoment
20            DistanceVector(1, 3) = Moments(3, 1) - m_Target3rdRawMoment

21            If (Abs(DistanceVector(1, 1)) + Abs(DistanceVector(1, 2)) + Abs(DistanceVector(1, 3))) < 0.000000001 Then
22                sSLNCalibrateNewtonRaphson = AppendSkewSign(GuessVec)
23                Exit Function
24            End If

25            AdjustmentVector = Application.WorksheetFunction.MMult(DistanceVector, inverseJacobian)

26            NextGuess(1) = GuessVec(1) - AdjustmentVector(1)
27            NextGuess(2) = GuessVec(2) - AdjustmentVector(2)
28            NextGuess(3) = GuessVec(3) - AdjustmentVector(3)

29            GuessVec = NextGuess

30        Next i
31        Throw "Failed to converge after " + CStr(Max_Iterations) + " iterations"

32        Exit Function
ErrHandler:
33        sSLNCalibrateNewtonRaphson = "#sSLNCalibrateNewtonRaphson (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNCalibrateMinPack
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Calibrates a shifted log-normal distribution to a basket of assets with given
'             forwards, weights, vols and correlation structure. Uses MinPack solver.
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNCalibrateMinPack(ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double)
1         On Error GoTo ErrHandler
          Dim ChooseVector
          Dim GuessVec
          Dim NumNonZeroWeights
          Dim numWeights As Long
          Dim Res

2         ChooseVector = sArrayNot(sArrayEquals(Weights, 0))
3         numWeights = sNRows(Weights)
4         NumNonZeroWeights = sArrayCount(ChooseVector)

5         If NumNonZeroWeights = 0 Then
6             sSLNCalibrateMinPack = sArrayStack(0, -200, 0, False)
7             Exit Function
8         End If

9         If NumNonZeroWeights < numWeights Then
10            Forwards = sMChoose(Forwards, ChooseVector)
11            Weights = sMChoose(Weights, ChooseVector)
12            Vols = sMChoose(Vols, ChooseVector)
13            Correlations = sMChoose(Correlations, ChooseVector)
              'Have no RowMChoose function so have to Transpose twice.
14            Correlations = sArrayTranspose(sMChoose(sArrayTranspose(Correlations), ChooseVector))
15        End If

16        If NumNonZeroWeights = 1 Then        'no calibration necessary
17            If Weights(1, 1) < 0 Then
18                sSLNCalibrateMinPack = sArrayStack(0, _
                      Log(Forwards(1, 1) * -Weights(1, 1)) - 0.5 * Vols(1, 1) ^ 2 * Time, _
                      Vols(1, 1) * Sqr(Time), _
                      True)
19            Else

20                sSLNCalibrateMinPack = sArrayStack(0, _
                      Log(Forwards(1, 1) * Weights(1, 1)) - 0.5 * Vols(1, 1) ^ 2 * Time, _
                      Vols(1, 1) * Sqr(Time), _
                      False)

21            End If
22            Exit Function
23        End If
24        GuessVec = GuessFirstSolution(Forwards, Weights, Vols, Correlations, Time)
25        Res = fsolve("BasketSolveObjectiveFn", GuessVec)
26        sSLNCalibrateMinPack = AppendSkewSign(Res)

27        Exit Function
ErrHandler:
28        sSLNCalibrateMinPack = "#sSLNCalibrateMinPack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNRawMoments
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Returns 3x1 vector of raw moments (i.e. moments about zero) of the shifted log nomal distribution.
'            M1 = Tau + exp(m + 1/2 s^2)
'            M2 = Tau^2 + 2.Tau.exp(m + 1/2 s^2) + exp(2m + 2s^2)
'            M3 = Tau^3 + 3.Tau^2.exp(m + 1/2 s^2) + 3.Tau.exp(2m + 2s^2) + exp(3m + 9/2 (s^2))
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNRawMoments(Tau As Double, M As Double, S As Double)
Attribute sSLNRawMoments.VB_Description = "First three moments of a shifted log normal distribution with given parameters. Tau is the shift, m and s are the mean and standard deviation of the underlying normal distribution."
Attribute sSLNRawMoments.VB_ProcData.VB_Invoke_Func = " \n29"
          Dim exp1 As Double
          Dim exp2 As Double
          Dim exp3 As Double
          Dim Result()

1         On Error GoTo ErrHandler
2         exp1 = Exp(M + 0.5 * S ^ 2)
3         exp2 = Exp(2 * M + 2 * S ^ 2)
4         exp3 = Exp(3 * M + 9 / 2 * S ^ 2)

5         ReDim Result(1 To 3, 1 To 1)

6         Result(1, 1) = Tau + exp1
7         Result(2, 1) = Tau ^ 2 + 2 * Tau * exp1 + exp2
8         Result(3, 1) = Tau ^ 3 + 3 * Tau ^ 2 * exp1 + 3 * Tau * exp2 + exp3

9         sSLNRawMoments = Result
10        Exit Function
ErrHandler:
11        Throw "#sSLNRawMoments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNCentralMoments
' Author    : Philip Swannell
' Date      : 09-May-2015
' Purpose   : Returns the Mean (1st raw moment) plus the second and third central moments
'             of a shifted log normal distribution with parameters Tau, m, s
'             M1 = Tau + exp(m + 1/2 s^2)
'             M'2 = exp(2m + 2s^2)- exp(2m+s^2)
'             M'3 = exp (3m + 9/2 s^s) -3exp(3m+ 5/2 s^2) + 2 exp(3m + 3/2 s^2)
'             as check we should have (for any distribution, central moments in terms of raw moments)
'             M'2 = M2 - M1^2
'             M'3 = M3 - 3M1M2 +2M1^3
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNCentralMoments(Tau As Double, M As Double, S As Double)
          Dim Result(1 To 3, 1 To 1) As Double

1         On Error GoTo ErrHandler

2         Result(1, 1) = Tau + Exp(M + 0.5 * S ^ 2)
3         Result(2, 1) = Exp(2 * M + 2 * S ^ 2) - Exp(2 * M + S ^ 2)
4         Result(3, 1) = Exp(3 * M + 9 / 2 * S ^ 2) - 3 * Exp(3 * M + 5 / 2 * S ^ 2) + 2 * Exp(3 * M + 3 / 2 * S ^ 2)

5         sSLNCentralMoments = Result
6         Exit Function
ErrHandler:
7         Throw "#sSLNCentralMoments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNJacobian
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Returns the Jacobian of the 3 raw moments as functions of Tau (the shift),
'             m (mean of normal distribution) and s (standard deviation of normal distribution)
'            Return is:
'            dM1/dTau    dM2/dTau    dM3/dTau
'            dM1/dm      dM2/dm      dM3/dm
'            dM1/ds      dM2/ds      dM3/ds
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNJacobian(Tau As Double, M As Double, S As Double)
Attribute sSLNJacobian.VB_Description = "Jacobian of the first three moments of a shifted log normal distribution with respect to its parameters. Tau is the shift, m and s are the mean and standard deviation of the underlying normal distribution. Hence:\nFirst moment = Tau + exp(m + 1/2 s^2)"
Attribute sSLNJacobian.VB_ProcData.VB_Invoke_Func = " \n29"
          Dim exp1 As Double
          Dim exp2 As Double
          Dim exp3 As Double
          Dim Result(1 To 3, 1 To 3)

1         On Error GoTo ErrHandler
2         exp1 = Exp(M + 0.5 * S ^ 2)
3         exp2 = Exp(2 * M + 2 * S ^ 2)
4         exp3 = Exp(3 * M + 9 / 2 * S ^ 2)

5         Result(1, 1) = 1        'dM1/dTau
6         Result(2, 1) = exp1        'dM1/dm
7         Result(3, 1) = S * exp1        'dM1/ds

8         Result(1, 2) = 2 * Tau + 2 * exp1        'dM2/dTau
9         Result(2, 2) = 2 * Tau * exp1 + 2 * exp2        'dM2/dm
10        Result(3, 2) = 2 * Tau * S * exp1 + 4 * S * exp2        'dM2/ds

11        Result(1, 3) = 3 * Tau ^ 2 + 6 * Tau * exp1 + 3 * exp2        'dM3/dTau
12        Result(2, 3) = 3 * Tau ^ 2 * exp1 + 6 * Tau * exp2 + 3 * exp3        'dM3/dm
13        Result(3, 3) = 3 * Tau ^ 2 * S * exp1 + 12 * S * Tau * exp2 + 9 * S * exp3        'dM3/ds

14        sSLNJacobian = Result
15        Exit Function
ErrHandler:
16        Throw "#sSLNJacobian (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBlackScholes
' Author    : Philip Swannell
' Date      : 04-May-2015
' Purpose   : Returns undiscounted vanilla option price under a log-normal (Black-Scholes) model. For a
'             call option, the function calculates:
'
'             V = E[(S-K)+].
'
'             Which, by the Black-Scholes formula is:
'
'             V = F.N(z)-K.N(z-s), where s = Vol.Root(t), and z = log(F/K)/s + s/2.
' Arguments
' CallPut   : can take values: "C" or "Call", "P" or "Put", "B" or "Buy", "S" or "Sell", "UD" or "Up
'             Digital", "DD" or "Down Digital". Case insensitive and space characters are
'             ignored so "Up Digital" is equivalent to "updigital". Can be an array.
' Forward   : Forward price for the underlying random asset S. Can be an array.
' Strike    : Strike price K for the option. Must match the same units as Forward. Can be an array.
' Volatility: Annualised log-normal volatility of the asset S, quoted on an Act/356 basis. Enter 20% as
'             0.2. Can be an array.
' Time      : Time in years until the expiry of the option. Can be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sBlackScholes(CallPut, Forward, Strike, Volatility, Time)
Attribute sBlackScholes.VB_Description = "Returns undiscounted vanilla option price under a log-normal (Black-Scholes) model. For a call option, the function calculates:\n\nV = E[(S-K)+].\n\nWhich, by the Black-Scholes formula is:\n\nV = F.N(z)-K.N(z-s), where s = Vol.Root(t), and z = log(F/K)/s + s/2."
Attribute sBlackScholes.VB_ProcData.VB_Invoke_Func = " \n29"

1         On Error GoTo ErrHandler

2         If VarType(CallPut) < vbArray And VarType(Forward) < vbArray And VarType(Strike) < vbArray And VarType(Volatility) < vbArray And VarType(Time) < vbArray Then
3             sBlackScholes = ThrowIfError(bsCore(StringsToOptStyle(CallPut, False), Forward, Strike, Volatility, Time))
4         Else
5             sBlackScholes = ThrowIfError(Broadcast(FuncIdBlackScholes, CallPut, Forward, Strike, Volatility, Time))
6         End If

7         Exit Function
ErrHandler:
8         sBlackScholes = "#sBlackScholes (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sOptSolveVol
' Author    : Philip Swannell
' Date      : 27-Nov-2015
' Purpose   : Returns the implied annualised volatility of an option with a given price, either under a
'             log-normal model or a normal model.
' Arguments
' CallPut   : Can take values: "C" or "Call", "P" or "Put", "B" or "Buy", "S" or "Sell", "UD" or "Up
'             Digital", "DD" or "Down Digital". Case insensitive and space characters are
'             ignored so "Up Digital" is equivalent to "updigital". Can be an array.
' Value     : Value of the option. Must match the same units as Forward and Strike. Can be an array.
' Forward   : Forward price for the underlying random asset S. Can be an array.
' Strike    : Strike price K for the option. Must match the same units as Forward. Can be an array.
' Time      : Time in years until the expiry of the option. Can be an array.
' LogNormal : Enter TRUE for a log-normal vol, or FALSE for a normal vol.
' -----------------------------------------------------------------------------------------------------------------------
Function sOptSolveVol(CallPut, Value, Forward, Strike, Time, LogNormal)
Attribute sOptSolveVol.VB_Description = "Returns the implied annualised volatility of an option with a given price, either under a log-normal model or a normal model."
Attribute sOptSolveVol.VB_ProcData.VB_Invoke_Func = " \n29"
1         On Error GoTo ErrHandler
2         If VarType(CallPut) < vbArray And VarType(Forward) < vbArray And VarType(Strike) < vbArray And VarType(Value) < vbArray And VarType(Time) < vbArray And VarType(LogNormal) < vbArray Then
3             sOptSolveVol = ThrowIfError(CoreOptSolveVol(StringsToOptStyle(CallPut, False), Value, Forward, Strike, Time, LogNormal))
4         Else
5             sOptSolveVol = ThrowIfError(Broadcast(FuncIdOptSolveVol, CallPut, Value, Forward, Strike, Time, LogNormal))
6         End If
7         Exit Function
ErrHandler:
8         sOptSolveVol = "#sOptSolveVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNormOpt
' Author    : Philip Swannell
' Date      : 27-Nov-2015
' Purpose   : Undiscounted vanilla option price under a normal model. For a call option, the function
'             calculates:
'
'             V = E[(S-K)+].
'
'             Which is:
'
'             V = (F-K).N(z) + s.N'(-z), where s = Vol.Root(t), and z = (F-K)/s, N is
'             normal distribution fn and N' is normal density fn.
' Arguments
' CallPut   : can take values: "C" or "Call", "P" or "Put", "B" or "Buy", "S" or "Sell", "UD" or "Up
'             Digital", "DD" or "Down Digital". Case insensitive and space characters are
'             ignored so "Up Digital" is equivalent to "updigital". Can be an array.
' Forward   : Forward price for the underlying random asset S. Can be an array.
' Strike    : Strike price K for the option. Must match the same units as Forward. Can be an array.
' Volatility: Annualised normal volatility of the asset S, quoted on an Act/356 basis. Enter 20% as 0.2.
'             Can be an array.
' Time      :  Time in years until the expiry of the option. Can be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sNormOpt(CallPut, Forward, Strike, Volatility, Time)
Attribute sNormOpt.VB_Description = "Undiscounted vanilla option price under a normal model. For a call option, the function calculates:\n\nV = E[(S-K)+].\n\nWhich is:\n\nV = (F-K).N(z) + s.N'(-z), where s = Vol.Root(t), and z = (F-K)/s, N is normal distribution fn and N' is normal density fn."
Attribute sNormOpt.VB_ProcData.VB_Invoke_Func = " \n29"

1         On Error GoTo ErrHandler

2         If VarType(CallPut) < vbArray And VarType(Forward) < vbArray And VarType(Strike) < vbArray And VarType(Volatility) < vbArray And VarType(Time) < vbArray Then
3             sNormOpt = ThrowIfError(CoreNormOpt(StringsToOptStyle(CallPut, False), Forward, Strike, Volatility, Time))
4         Else
5             sNormOpt = ThrowIfError(Broadcast(FuncIdNormOpt, CallPut, Forward, Strike, Volatility, Time))
6         End If

7         Exit Function
ErrHandler:
8         sNormOpt = "#sNormOpt (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSLNEuropeanOption
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Price European options (Call, Put, Digital Call, Digital Put) against a shifted log normal
'             distribution. Arguments
'             Tau, m and s are parameters such that SLN = Tau + exp(N) where N is normal
'             with mean m and standard deviation s.
' -----------------------------------------------------------------------------------------------------------------------
Function sSLNEuropeanOption(Strike As Double, ByVal CallPut As String, Tau As Double, M As Double, S As Double, UseNegativeDistribution As Boolean)
Attribute sSLNEuropeanOption.VB_Description = "Returns the undiscounted value of a European option or digital option on an asset whose risk-neutral distribution is a shifted log-normal or shifted negative log-normal distribution."
Attribute sSLNEuropeanOption.VB_ProcData.VB_Invoke_Func = " \n29"
          Dim AdjustedStrike As Double
          Dim Forward As Double
          Dim OptionStyle As EnmOptStyle

1         On Error GoTo ErrHandler
2         AdjustedStrike = Strike - Tau
3         Forward = Exp(M + 0.5 * S ^ 2)
4         OptionStyle = StringsToOptStyle(CallPut, False)

5         If UseNegativeDistribution Then
6             AdjustedStrike = -Strike - Tau
              'Flip Call to Put and Put to Call
7             Select Case OptionStyle
                  Case OptStyleCall
8                     OptionStyle = OptStylePut
9                 Case OptStylePut
10                    OptionStyle = OptStyleCall
11                Case optStyleDownDigital
12                    OptionStyle = optStyleUpDigital
13                Case optStyleUpDigital
14                    OptionStyle = optStyleDownDigital
15                Case OptStyleBuy
16                    OptionStyle = OptStyleSell
17                Case OptStyleSell
18                    OptionStyle = OptStyleBuy
19                Case Else
20                    Throw "Unhandled OptionStyle"
21            End Select
22        End If
23        sSLNEuropeanOption = bsCore(OptionStyle, Forward, AdjustedStrike, S, 1)

24        Exit Function
ErrHandler:
25        sSLNEuropeanOption = "#sSLNEuropeanOption (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BasketSolveObjectiveFn
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Used by the MinPack fsolve function, which seeks to find Tau, m and s (shift, mean
'             and standard deviation) such that the first three moments of the
'            (possibly negative) shifted log normal match the moments of the basket
' -----------------------------------------------------------------------------------------------------------------------
Function BasketSolveObjectiveFn(Tau_m_s_Vector As Variant) As Double()
          Dim Moments
          Dim Result() As Double
1         On Error GoTo ErrHandler
2         ReDim Result(1 To 3) As Double

3         Moments = sSLNRawMoments(CDbl(Tau_m_s_Vector(1)), CDbl(Tau_m_s_Vector(2)), CDbl(Tau_m_s_Vector(3)))

4         Result(1) = (Moments(1, 1) - m_Target1stMoment) * 1000
5         Result(2) = (Moments(2, 1) - m_Target2ndRawMoment) * 1000
6         Result(3) = (Moments(3, 1) - m_Target3rdRawMoment) * 1000

7         BasketSolveObjectiveFn = Result
8         Exit Function
ErrHandler:
9         Throw "#BasketSolveObjectiveFn (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GuessFirstSolution
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Sets the 4 module-level variables (3 moments + m_BasketSkewIsNegative),
'             and calculates the first guess for the solution.
'             First guess matches the second and third moments with the shift set to zero.
'             Unfortuanely such a first-guess solution may not exist. in which case first guess is (0,0,0)
'             and the chance of either solver finding a solution is presumably reduced.
'             Is there a better way to do the first guess?
' -----------------------------------------------------------------------------------------------------------------------
Private Function GuessFirstSolution(Forwards, Weights, Vols, Correlations, Time)
1         On Error GoTo ErrHandler
          Dim GuessVec()
2         ReDim GuessVec(1 To 3)
          Dim Basket_Mean
          Dim Basket_Skew
          Dim Basket_Variance
          Dim NumAssets As Long

          'Dim Guess_Tau, Guess_m, Guess_s
3         NumAssets = UBound(Forwards, 1) - LBound(Forwards, 1) + 1

          'm_TargetMoments = sBasketRawMoments(Forwards, Weights, Vols, Correlations, Time)
4         m_Target1stMoment = sBasket1stMoment(Forwards, Weights, NumAssets)
5         m_Target2ndRawMoment = sBasket2ndRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)
6         m_Target3rdRawMoment = sBasket3rdRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)

7         Basket_Mean = m_Target1stMoment
8         Basket_Variance = m_Target2ndRawMoment - m_Target1stMoment ^ 2
9         Basket_Skew = (m_Target3rdRawMoment - 3 * m_Target2ndRawMoment * m_Target1stMoment + 2 * m_Target1stMoment ^ 3) / (Basket_Variance ^ 1.5)
10        m_BasketSkewIsNegative = (Basket_Skew < 0)

11        If m_BasketSkewIsNegative Then
12            m_Target1stMoment = -m_Target1stMoment
13            m_Target3rdRawMoment = -m_Target3rdRawMoment
14        End If

15        On Error Resume Next
16        GuessVec(1) = 0
17        GuessVec(3) = Sqr(2 / 3 * Log(m_Target3rdRawMoment) - Log(m_Target2ndRawMoment))
18        GuessVec(2) = 1 / 2 * Log(m_Target2ndRawMoment) - GuessVec(3) ^ 2
19        On Error GoTo ErrHandler

20        GuessFirstSolution = GuessVec

21        Exit Function
ErrHandler:
22        Throw "#GuessFirstSolution (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketRawMoments
' Author    : Philip Swannell
' Date      : 29-Apr-2015
' Purpose   : Returns the first three raw moments of the distribution at time T of the weighted sum of a
'             number of correlated log-normal assets as a column vector, i.e. E(X). E(X^2)
'             and E(X^3)
' Arguments
' Forwards  : A column vector of the forward prices of each asset.
' Weights   : A column vector of the weights in the basket of each asset.
' Vols      : A column vector of the log-normal volatilities of the assets. Enter 20% volatility as 0.2.
' Correlations: The correlations between the assets as a matrix. The function checks that the matrix is
'             symmetric, has on-diagonal values of 1 and no values outside the range
'             [-1,1]. It does not check that the matrix is positive definite.
' Time      : The time in years until the basket price is observed.
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketRawMoments(Forwards, Weights, Vols, Correlations, Time)
Attribute sBasketRawMoments.VB_Description = "Returns the first three raw moments of the distribution at time T of the weighted sum of a number of correlated log-normal assets as a column vector, i.e. E(X). E(X^2) and E(X^3)"
Attribute sBasketRawMoments.VB_ProcData.VB_Invoke_Func = " \n29"
          Dim NumAssets As Long
          Dim Result() As Variant

          'Should check inputs are well formed here...
1         On Error GoTo ErrHandler
2         Force2DArrayR Forwards: Force2DArrayR Weights: Force2DArrayR Vols: Force2DArrayR Correlations

3         ReDim Result(1 To 3, 1 To 1)
4         NumAssets = sNRows(Forwards)

5         Result(1, 1) = sBasket1stMoment(Forwards, Weights, NumAssets)
6         Result(2, 1) = sBasket2ndRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)
7         Result(3, 1) = sBasket3rdRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)

8         sBasketRawMoments = Result
9         Exit Function
ErrHandler:
10        Throw "#sBasketRawMoments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketCentralMoments
' Author    : Philip Swannell
' Date      : 09-May-2015
' Purpose   : Returns a vector of Mean, E((X-Mean)^2) and E((X-Mean)^3)
'             E(X)
'             E((X-Mean)^2) = E(X^2) - (E(X))^2
'             E((X-Mean)^3) = E(X^3) - 3E(X^2)E(X) + 2(E(X))^3
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketCentralMoments(Forwards, Weights, Vols, Correlations, Time)
1         On Error GoTo ErrHandler
          Dim M1 As Double
          Dim M2 As Double
          Dim M3 As Double
          Dim NumAssets As Long
          Dim Result(1 To 3, 1 To 1)

2         CheckBasketInputs Forwards, Weights, Vols, Correlations
3         NumAssets = sNRows(Forwards)
4         M1 = sBasket1stMoment(Forwards, Weights, NumAssets)
5         M2 = sBasket2ndRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)
6         M3 = sBasket3rdRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)

7         Result(1, 1) = M1
8         Result(2, 1) = M2 - M1 ^ 2
9         Result(3, 1) = M3 - 3 * M1 * M2 + 2 * M1 ^ 3
10        sBasketCentralMoments = Result

11        Exit Function
ErrHandler:
12        Throw "#sBasketCentralMoments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasket1stMoment
' Author    : Philip Swannell
' Date      : 29-Apr-2015
' Purpose   : Returns 1st moment of basket. See page 3 of "American Basket and Spread Option
'             pricing by a Simple Binomial Tree" by S. Borokova et al
' -----------------------------------------------------------------------------------------------------------------------
Private Function sBasket1stMoment(Forwards, Weights, NumAssets)
          Dim i As Long
          Dim Result As Double
1         On Error GoTo ErrHandler

          'Assumes inputs are well formed - all numbers arrays of correct size etc
2         For i = 1 To NumAssets
3             Result = Result + Weights(i, 1) * Forwards(i, 1)
4         Next i
5         sBasket1stMoment = Result

6         Exit Function
ErrHandler:
7         Throw "#sBasket1stMoment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasket2ndRawMoment
' Author    : Philip Swannell
' Date      : 29-Apr-2015
' Purpose   : Returns 2nd moment of basket. See page 3 of "American Basket and Spread Option
'             pricing by a Simple Binomial Tree" by S. Borokova et al
' -----------------------------------------------------------------------------------------------------------------------
Private Function sBasket2ndRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets)
          Dim i As Long
          Dim j As Long
          Dim Result As Double
1         On Error GoTo ErrHandler

2         For i = 1 To NumAssets
3             For j = 1 To NumAssets
4                 Result = Result + Weights(i, 1) * Weights(j, 1) * Forwards(i, 1) * Forwards(j, 1) * Exp(Correlations(i, j) * Vols(i, 1) * Vols(j, 1) * Time)
5             Next j
6         Next i

7         sBasket2ndRawMoment = Result

8         Exit Function
ErrHandler:
9         Throw "#sBasket2ndRawMoment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasket3rdRawMoment
' Author    : Philip Swannell
' Date      : 29-Apr-2015
' Purpose   : Returns 3rd moment of basket. See page 3 of "American Basket and Spread Option
'             pricing by a Simple Binomial Tree" by S. Borokova et al
' -----------------------------------------------------------------------------------------------------------------------
Private Function sBasket3rdRawMoment(Forwards, Weights, Vols, Correlations, Time, NumAssets As Long)

          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim Result As Double

1         On Error GoTo ErrHandler

          'Assumes inputs are well formed - all numbers arrays of correct size etc

2         For i = 1 To NumAssets
3             For j = 1 To NumAssets
4                 For k = 1 To NumAssets
5                     Result = Result + Weights(i, 1) * Weights(j, 1) * Weights(k, 1) * Forwards(i, 1) * Forwards(j, 1) * Forwards(k, 1) * _
                          Exp((Correlations(i, j) * Vols(i, 1) * Vols(j, 1) + Correlations(i, k) * Vols(i, 1) * Vols(k, 1) + Correlations(j, k) * Vols(j, 1) * Vols(k, 1)) * Time)
6                 Next k
7             Next j
8         Next i

9         sBasket3rdRawMoment = Result

10        Exit Function
ErrHandler:
11        Throw "#sBasket3rdRawMoment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AppendSkewSign
' Author    : Philip Swannell
' Date      : 01-May-2015
' Purpose   : Any caller of the calibration process needs to know the sign of the skew since
'             that determines if we fit a shifted log-normal or a shifted negative log-normal.
'             This utility function tacks that information on.
' -----------------------------------------------------------------------------------------------------------------------
Private Function AppendSkewSign(Tau_m_s_Vector)
          Dim Result
1         On Error GoTo ErrHandler
2         ReDim Result(1 To 4, 1 To 1)
3         Result(1, 1) = Tau_m_s_Vector(1)
4         Result(2, 1) = Tau_m_s_Vector(2)
5         Result(3, 1) = Tau_m_s_Vector(3)
6         Result(4, 1) = m_BasketSkewIsNegative
7         AppendSkewSign = Result
8         Exit Function
ErrHandler:
9         Throw "#AppendSkewSign (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFxForwardPFE
' Author    : Philip Swannell
' Date      : 11-May-2015
' Purpose   : A "mid-level" function for calculating the PFE of a portfolio of Fx forwards. Wraps sSLNCalibrateMinPack (the low-level function)
'             and itself needs wrapping with a layer that interpolates vols, forwards and holdings at the observation dates calculated from trade data.
' Arguments:
'  AnchorDate: The "Today" date - so discount factors are one for this date
'  DomCcy:     The domestic currency as a three letter swift code.
'  FgnCcys:    A column array of foreign currencies as 3 letter codes. Must all be different from DomCCy
'  Forwards:   a 2-d array. Same no. of cols as ObservationDates, same no of rows as FgnCcys. Each element
'              gives the Fgn/Dom forward Fx rate for the corresponding FgnCcy and ObservationDate
'  FgnAmounts: 2-d array same size as Forwards. Each element gives the amount of corresponding foreign currency
'              held as of corresponding observation date.
'  Vols:        2-d array same size as Forwards. Each element gives term vol of Fgn/Dom FxRate to corresponding observation date.
'  Correlations: A symmetric matrix giving the correlations between the Fgn/Dom Fx rates. Hence no term structure for correlation.
'  ObservationDates: a row vector of observation dates, on each of which the function calculates PFE at the given confidence levels.
'                    ObservationDates need not be whole numbers
'  ConfidenceLevels: Column array of confidence levels eg {0.99;0.95}
'  ShowDetails: Boolean, If TRUE then the return has more rows giving forwards, vols, amounts at each of the observation dates.
' -----------------------------------------------------------------------------------------------------------------------
Function sFxForwardPFE(AnchorDate, DomCCy, FgnCcys, DomAmounts, Forwards, FgnAmounts, Vols, Correlations, ObservationDates, ConfidenceLevels, ShowDetails As Boolean, CalcMethod As String, NumMCPaths As Long)
          Dim CurrentPV
          Dim i As Long
          Dim j As Long
          Dim NumConfidenceLevels As Long
          Dim NumTimes As Long
          Dim PFERes
          Dim Result() As Variant
          Dim ThisDomAmount
          Dim ThisFgnAmounts
          Dim ThisForwards
          Dim ThisTime
          Dim ThisVols
          Dim TimeVector
          Dim UseMomentMatching As Boolean

1         On Error GoTo ErrHandler

2         Select Case LCase$(Replace(CalcMethod, " ", vbNullString))
              Case "momentmatching", "moment-matching"
3                 UseMomentMatching = True
4             Case "sobol", "shifted-sobol", "wichmann-hill", "wichmannhill", "mersennetwister", "mersenne-twister", "vbarnd", "vba-rnd"
5             Case Else
6                 Throw "Unrecognised CalcMethod - allowed are: Moment-Matching, Sobol, Wichmann-Hill, Mersenne-Twister, VBA-Rnd and Shifted-Sobol"
7         End Select

8         Force2DArrayRMulti FgnCcys, DomAmounts, Forwards, FgnAmounts, Vols, Correlations, ObservationDates, ConfidenceLevels
          'Check all inputs here!
9         TimeVector = sArrayDivide(sArraySubtract(ObservationDates, AnchorDate), 365)
10        NumTimes = sNCols(ObservationDates)
11        NumConfidenceLevels = sNRows(ConfidenceLevels)

12        ReDim Result(1 To NumConfidenceLevels + 1, 1 To NumTimes + 1)
          'Populate headers...
13        Result(1, 1) = "Time"
14        For i = 1 To NumConfidenceLevels
15            Result(i + 1, 1) = Format$(ConfidenceLevels(i, 1), "0.0%") + " Confidence PFE"
16        Next i
17        For j = 1 To NumTimes
18            Result(1, j + 1) = TimeVector(1, j)
19        Next

20        If Not UseMomentMatching Then
              Dim Cholesky
              Dim CorrelatedNormals
              Dim Normals
21            Cholesky = ThrowIfError(sCholesky(Correlations))
22            Normals = ThrowIfError(sRandomVariable(NumMCPaths, sNRows(Forwards), "Normal", CalcMethod))
23            CorrelatedNormals = Application.WorksheetFunction.MMult(Normals, sArrayTranspose(Cholesky))
24        End If

25        For j = 1 To NumTimes
26            ThisTime = TimeVector(1, j)
27            ThisDomAmount = DomAmounts(1, j)
28            ThisForwards = sSubArray(Forwards, , j, , 1)
29            ThisFgnAmounts = sSubArray(FgnAmounts, , j, , 1)
30            ThisVols = sSubArray(Vols, , j, , 1)

31            If ThisTime < 0 Then
32                For i = 1 To NumConfidenceLevels
33                    Result(i + 1, j + 1) = "#Time is negative!"
34                Next i
35            ElseIf ThisTime = 0 Then
36                CurrentPV = ThisDomAmount + sColumnSum(sArrayMultiply(ThisForwards, ThisFgnAmounts))(1, 1)
37                For i = 1 To NumConfidenceLevels
38                    Result(i + 1, j + 1) = CurrentPV
39                Next i
40            Else
41                If UseMomentMatching Then
42                    PFERes = sBasketPFE(ThisForwards, ThisFgnAmounts, ThisVols, Correlations, CDbl(ThisTime), ConfidenceLevels)
43                Else
44                    PFERes = sBasketPFEMC(ThisForwards, ThisFgnAmounts, ThisVols, Correlations, CDbl(ThisTime), NumMCPaths, CalcMethod, ConfidenceLevels, CorrelatedNormals)
45                End If
46                For i = 1 To NumConfidenceLevels
47                    If VarType(PFERes(i, 1)) = vbString Then
48                        Result(i + 1, j + 1) = PFERes(i, 1)
49                    Else
50                        Result(i + 1, j + 1) = ThisDomAmount + PFERes(i, 1)
51                    End If
52                Next i
53            End If
54        Next j

55        If ShowDetails Then
56            Result = sArrayStack(sArrayRange("Date", ObservationDates), Result, _
                  sArrayRange(sArrayConcatenate(FgnCcys, "/", DomCCy, " Forward"), Forwards), _
                  sArrayRange(sArrayConcatenate(FgnCcys, "/", DomCCy, " Vol"), Vols), _
                  sArrayRange(DomCCy & " BasketWeight", DomAmounts), _
                  sArrayRange(sArrayConcatenate(FgnCcys, " BasketWeight"), FgnAmounts))
57        Else
58            Result = sArrayStack(sArrayRange("Date", ObservationDates), Result)
59        End If

60        sFxForwardPFE = Result

61        Exit Function
ErrHandler:
62        sFxForwardPFE = "#sFxForwardPFE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketPFEMC
' Author    : Philip Swannell
' Date      : 06-Jun-2015
' Purpose   : Uses Monte Carlo to calculate Potential Future Exposure on a basket of correlated
'             log normal assets at one or more confidence levels.
'             CorrelatedNormals is optional. If passed (for speed of repeated calls) it should have NumPaths rows, same number of columns as Forwards has rows,
'             each column should be a sample from N(0,1) and the correlation structure should be as given by Correlations.
'             If not passed the CorrelatedNormals is calculated from the given RNGName (see sRandomVariable) and Correlation
' Potential speedup - detect when there is only one non-zero weight and give analytic result. Two non-zero weights might also be tractable
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketPFEMC(ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double, NumPaths As Long, RNGName As String, ConfidenceLevels, Optional CorrelatedNormals As Variant)

1         On Error GoTo ErrHandler

          Dim BasketPaths As Variant
          Dim Cholesky As Variant
          Dim CurrentMeans As Variant
          Dim CurrentStDev As Variant
          Dim DesiredMeans As Variant
          Dim DesiredStDev As Variant
          Dim LogNormals As Variant
          Dim Normals As Variant
          Dim NumAssets As Long
          Dim NumNonZeroWeights
          Dim Rebase As Boolean
          Dim RebasedCorrelatedNormals

2         CheckBasketInputs Forwards, Weights, Vols, Correlations
3         Force2DArrayR ConfidenceLevels
4         If NumPaths < 1 Then Throw "NumPaths must be positive"
5         NumAssets = sNRows(Forwards)
6         NumNonZeroWeights = NumAssets - sArrayCount(sArrayEquals(Weights, 0))
7         If NumNonZeroWeights = 0 Then        'Could avoid Monte Carlo altogether if NumNonZeroWeights is 1!
8             sBasketPFEMC = sReshape(0, sNRows(ConfidenceLevels), 1)
9             Exit Function
10        End If

11        If IsMissing(CorrelatedNormals) Then
12            Normals = ThrowIfError(sRandomVariable(NumPaths, NumAssets, "Normal", RNGName))
13            If sArraysIdentical(Correlations, sIdentityMatrix(sNRows(Forwards))) Then
14                CorrelatedNormals = Normals
15            Else
16                Cholesky = ThrowIfError(sCholesky(Correlations))
17                CorrelatedNormals = Application.WorksheetFunction.MMult(Normals, sArrayTranspose(Cholesky))
18            End If
19        End If

20        Rebase = InStr(LCase$(RNGName), "sobol") = 0        'We don't correct mean or standard dev for sobol sequences

21        DesiredMeans = sArrayTranspose(sArraySubtract(sArrayLog(Forwards), sArrayMultiply(0.5, Vols, Vols, Time)))
22        DesiredStDev = sArrayTranspose(sArrayMultiply(Vols, sArrayPower(Time, 0.5)))
23        If Rebase Then
24            CurrentMeans = sColumnMean(CorrelatedNormals)
25            CurrentStDev = sColumnStDev(CorrelatedNormals)
26            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(sArraySubtract(CorrelatedNormals, CurrentMeans), sArrayDivide(DesiredStDev, CurrentStDev)), DesiredMeans)
27        Else
28            RebasedCorrelatedNormals = sArrayAdd(sArrayMultiply(CorrelatedNormals, DesiredStDev), DesiredMeans)
29        End If

30        LogNormals = sArrayExp(RebasedCorrelatedNormals)
31        BasketPaths = Application.WorksheetFunction.MMult(LogNormals, Weights)
32        sBasketPFEMC = ThrowIfError(sQuantiles(BasketPaths, ConfidenceLevels))

33        Exit Function
ErrHandler:
34        Throw "#sBasketPFEMC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIdentityMatrix
' Author    : Philip Swannell
' Date      : 07-Jun-2015
' Purpose   : Returns the identity matrix of dimension N. On diagonal elements are 1, off diagonal
'             elements are 0. This function duplicates the Excel function MUNIT.
' Arguments
' N         : The number of rows and columns in the return.
' -----------------------------------------------------------------------------------------------------------------------
Function sIdentityMatrix(N As Variant)
Attribute sIdentityMatrix.VB_Description = "Returns the identity matrix of dimension N. On diagonal elements are 1, off diagonal elements are 0. This function duplicates the Excel function MUNIT."
Attribute sIdentityMatrix.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim R() As Double
1         On Error GoTo ErrHandler
2         ReDim R(1 To N, 1 To N)
3         For i = 1 To N
4             R(i, i) = 1
5         Next i
6         sIdentityMatrix = R
7         Exit Function
ErrHandler:
8         Throw "#sIdentityMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDiagonal
' Author    : Philip Swannell
' Date      : 28-Jun-2017
' Purpose   : When the input is square, returns a vector of the diagonal elements. When the input has
'             one row or one column, returns a square matrix whose diagonal is the input
'             and off-diagonal is zero.
' Arguments
' VectorOrMatrix: Either a square matrix or an array with one column or one row.
' -----------------------------------------------------------------------------------------------------------------------
Function sDiagonal(ByVal VectorOrMatrix As Variant)
Attribute sDiagonal.VB_Description = "When the input is square, returns a vector of the diagonal elements. When the input has one row or one column, returns a square matrix whose diagonal is the input and off-diagonal is zero."
Attribute sDiagonal.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res
1         On Error GoTo ErrHandler
2         Force2DArrayR VectorOrMatrix, NR, NC

3         If NR = NC Then
4             Res = sReshape(0, NR, 1)
5             For i = 1 To NR
6                 Res(i, 1) = VectorOrMatrix(i, i)
7             Next i
8         ElseIf NR = 1 Then
9             Res = sReshape(0, NC, NC)
10            For i = 1 To NC
11                Res(i, i) = VectorOrMatrix(1, i)
12            Next i
13        ElseIf NC = 1 Then
14            Res = sReshape(0, NR, NR)
15            For i = 1 To NR
16                Res(i, i) = VectorOrMatrix(i, 1)
17            Next i
18        Else
19            Throw "VectorOrMatrix must either have the same number of rows as columns, or have one row or have one column"
20        End If
21        sDiagonal = Res

22        Exit Function
ErrHandler:
23        sDiagonal = "#sDiagonal (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestQuantiles2
' Author    : Philip Swannell
' Date      : 22-Jul-2017
' Purpose   : Test bench function, optimised for speed by minimising the number of calls to sRandomVariable
'             SampleSizes should be a column array, QuantileTypes and Quantiles should be
'             1 - row arrays each with the same number of cols.
'             Return shows the expected errors in the differently-calculated quantiles (versus NormSInv)
'             and the standard errors of the results, and optionally the "Raw Data", which can then be analysed with sHistogram
' -----------------------------------------------------------------------------------------------------------------------
Private Function TestQuantiles2(SampleSizes As Variant, DistributionName As String, NumTests As Long, QuantileTypes As Variant, Quantiles As Variant, Seed As Long, RNGName As String, WithRawData As Boolean)

1         On Error GoTo ErrHandler
          Dim MaxSampleSize As Long
          Dim NC As Long
          Dim NR As Long
          Dim ThisQuantile As Double
          Dim ThisQuantileType As String
          Dim ThisSampleSize As Long

2         Force2DArrayRMulti SampleSizes, QuantileTypes, Quantiles

3         sRandomSetSeed RNGName, Seed

          Dim CutDownSample
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim QuantileRes
          Dim RunningSquaresTotal
          Dim RunningTotal() As Double
          Dim Sample

4         MaxSampleSize = sColumnMax(SampleSizes)(1, 1)
5         NC = sNCols(QuantileTypes)
6         NR = sNRows(SampleSizes)
7         ReDim RunningTotal(1 To NR, 1 To NC)
8         ReDim RunningSquaresTotal(1 To NR, 1 To NC)
          Dim ColNo As Long
9         If WithRawData Then
              Dim RawData As Variant
10            RawData = sReshape(0, NumTests + 1, NC)
11        End If
          Dim Corrects
12        Corrects = sReshape(0, 1, NC)
13        For j = 1 To NC
14            Corrects(1, j) = Application.WorksheetFunction.Norm_S_Inv(Quantiles(1, j))
15        Next j

16        For k = 1 To NumTests
17            Sample = ThrowIfError(sRandomVariable(MaxSampleSize, 1, DistributionName, RNGName))
18            ColNo = 0
19            For i = 1 To NR
20                ThisSampleSize = SampleSizes(i, 1)
21                If ThisSampleSize = MaxSampleSize Then
22                    CutDownSample = Sample
23                Else
24                    CutDownSample = sSubArray(Sample, 1, 1, ThisSampleSize, 1)
25                End If
26                For j = 1 To NC

27                    ThisQuantileType = QuantileTypes(1, j)
28                    ThisQuantile = Quantiles(1, j)
29                    QuantileRes = ThrowIfError(sGeneralisedQuantile2(CutDownSample, ThisQuantile, ThisQuantileType, False))
30                    RunningTotal(i, j) = RunningTotal(i, j) + QuantileRes
31                    RunningSquaresTotal(i, j) = RunningSquaresTotal(i, j) + QuantileRes ^ 2

32                    If WithRawData Then
33                        If ThisSampleSize = MaxSampleSize Then
34                            ColNo = ColNo + 1
35                            If k = 1 Then RawData(1, ColNo) = ThisQuantileType + " " + Format$(ThisQuantile, "0%")
36                            RawData(k + 1, ColNo) = QuantileRes - Corrects(1, ColNo)
37                        End If
38                    End If

39                Next j
40            Next i
41        Next k

          Dim Correct As Double
          Dim Result

42        Result = sArrayDivide(RunningTotal, NumTests)
43        For j = 1 To NC
44            Correct = Application.WorksheetFunction.Norm_S_Inv(Quantiles(1, j))
45            For i = 1 To NR
46                Result(i, j) = Result(i, j) - Correct
47            Next i
48        Next j

          Dim StandardErrors

49        StandardErrors = sArrayPower(sArraySubtract(sArrayDivide(RunningSquaresTotal, NumTests), sArrayPower(sArrayDivide(RunningTotal, NumTests), 2)), 0.5)

50        If WithRawData Then
51            TestQuantiles2 = sArrayStack(Result, StandardErrors, RawData)
52        Else
53            TestQuantiles2 = sArrayStack(Result, StandardErrors)
54        End If

55        Exit Function
ErrHandler:
56        TestQuantiles2 = "#TestQuantiles2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sGeneralisedQuantile
' Author    : Philip
' Date      : 20-Jul-2017
' Purpose   : Generalisation of PERCENTILE.EXC and PERCENTILE.INC.
'             x interpolation points are (m-c)/(n+1-2c) for m = 1, 2,...,n
'             y interpolation points are sorted Sample
' -----------------------------------------------------------------------------------------------------------------------
Function sGeneralisedQuantile(ByVal Sample, Quantile As Double, QuantileType As String, Optional IgnoreNonNumbers As Boolean)
1         On Error GoTo ErrHandler
          Dim c As Double
          Dim Denominator As Long
          Dim i As Long
          Dim j As Long
          Dim LowerM As Long
          Dim N As Long
          Dim NC As Long
          Dim NR As Long
          Dim UpperM As Long
          Dim x1 As Double
          Dim X2 As Double
          Dim y1 As Double
          Dim y2 As Double

2         Force2DArrayR Sample

3         Select Case UCase$(QuantileType)
              Case "INC"
4                 c = 1    'interp points are (m-1)/(n-1)
5             Case "EXC"
6                 c = 0    'interp points are m/(n+1)
7             Case "CENTRAL"
8                 c = 0.5    'interp points are (m-0.5)/n
9             Case Else
10                Throw "QuantileType must be 'INC', 'EXC' or 'CENTRAL'"
11        End Select

12        NR = sNRows(Sample): NC = sNCols(Sample)

13        For i = 1 To NR
14            For j = 1 To NC
15                If IsNumberOrDate(Sample(i, j)) Then
16                    N = N + 1
17                ElseIf IgnoreNonNumbers Then
18                    Sample(i, j) = vbNullString    ' will be treated correctly by call to Application.WorksheetFunction.Small
19                Else
20                    Throw "Non numbers encountered in Sample. Consider setting IgnoreNonNumbers to TRUE"
21                End If
22            Next j
23        Next i

24        If IgnoreNonNumbers Then
25            If N = 0 Then
26                Throw "No numbers encountered in Sample"
27            End If
28        End If

29        Denominator = N + 1 - 2 * c

30        If Quantile < (1 - c) / Denominator Or Quantile > (N - c) / Denominator Then
31            If UCase$(QuantileType) = "CENTRAL" Then
32                Throw "For QuantileType = " & UCase$(QuantileType) & " and Sample size = " + Format$(N, "###,##0") + ", Quantile must be in the range 1/" + Format$(2 * Denominator, "###,##0") + " to " + Format$(2 * N - 1, "###,##0") + "/" + Format$(2 * Denominator, "###,##0")
33            ElseIf UCase$(QuantileType) = "INC" Then
34                Throw "For QuantileType = " & UCase$(QuantileType) + ", Quantile must be in the range 0 to 1"
35            Else
36                Throw "For QuantileType = " & UCase$(QuantileType) & " and Sample size = " + Format$(N, "###,##0") + ", Quantile must be in the range " + CStr(1 - c) + "/" + Format$(Denominator, "###,##0") + " to " + Format$(N - c, "###,##0") + "/" + Format$(Denominator, "###,##0")
37            End If
38        End If

39        LowerM = RoundDown((Quantile * Denominator) + c)
40        UpperM = LowerM + 1

41        x1 = (LowerM - c) / Denominator
42        X2 = (UpperM - c) / Denominator

43        If Quantile < x1 Or Quantile > X2 Then Throw "Assertion failed: Quantile outside range x1 to x2"

44        y1 = Application.WorksheetFunction.Small(Sample, LowerM)
45        If Quantile > x1 Then
46            y2 = Application.WorksheetFunction.Small(Sample, UpperM)
47        Else
48            y2 = y1
49        End If

50        sGeneralisedQuantile = y1 + (Quantile - x1) / (X2 - x1) * (y2 - y1)
51        Exit Function
ErrHandler:
52        sGeneralisedQuantile = "#sGeneralisedQuantile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sGeneralisedQuantile2
' Author    : Philip
' Date      : 21-Jul-2017
' Purpose   : Implementation of Hyndman, R. J. and Fan, Y. (1996) Sample quantiles in statistical packages, American Statistician 50, 361–365.
'
'             x interpolation points are (m-c)/(n+1-2c) for m = 1, 2,...,n
'             y interpolation points are sorted Sample
' -----------------------------------------------------------------------------------------------------------------------
Function sGeneralisedQuantile2(ByVal Sample, p As Double, QuantileType As String, Optional IgnoreNonNumbers As Boolean)
1         On Error GoTo ErrHandler
          Dim g As Double
          Dim j As Long
          Dim M As Double

          Dim i As Long
          Dim k As Long
          Dim N As Long
          Dim NC As Long
          Dim NR As Long

2         Force2DArrayR Sample

3         Select Case UCase$(QuantileType)
              Case "4"
4                 M = 0
5             Case "CENTRAL", "5"
6                 M = 0.5
7             Case "EXC", "6"
8                 M = p
9             Case "INC", "7"
10                M = 1 - p
11            Case "8"
12                M = (p + 1) / 3
13            Case "9"
14                M = p / 4 + 3 / 8
15            Case Else
16                Throw "QuantileType not recognised. Allowed values are: '4', '5' (or 'CENTRAL'), '6' (or 'EXC'), 7 (or 'INC'), '8' ,'9'"
17        End Select

18        NR = sNRows(Sample): NC = sNCols(Sample)

19        For i = 1 To NR
20            For k = 1 To NC
21                If IsNumberOrDate(Sample(i, k)) Then
22                    N = N + 1
23                ElseIf IgnoreNonNumbers Then
24                    Sample(i, k) = vbNullString    ' will be treated correctly by call to Application.WorksheetFunction.Small
25                Else
26                    Throw "Non numbers encountered in Sample. Consider setting IgnoreNonNumbers to TRUE"
27                End If
28            Next k
29        Next i

30        If IgnoreNonNumbers Then
31            If N = 0 Then
32                Throw "No numbers encountered in Sample"
33            End If
34        End If

35        If p < (1 - M) / N Or p > (1 - M / N) Then
              Dim ErrorMessage As String
              Dim LB As String
              Dim UB As String
36            ErrorMessage = "For QuantileType = " & UCase$(QuantileType) & " and Sample size = " + Format$(N, "###,##0") + ", Quantile must be in the range "
37            LB = CStr((1 - M) / N)
38            UB = CStr(1 - M / N)
39            Select Case UCase$(QuantileType)
                  Case "4"
                      'TODO...

40                Case "CENTRAL", "5"
41                    LB = "1/" + Format$(2 * N, "###,##0")
42                    UB = Format$(2 * N - 1, "###,##0") + "/" + Format$(2 * N, "###,##0")
43                Case "EXC", "6"
44                    LB = CStr(1) + "/" + Format$(N + 1, "###,##0")
45                    UB = Format$(N, "###,##0") + "/" + Format$(N + 1, "###,##0")
46                Case "INC", "7"
47                    Throw "For QuantileType = " & UCase$(QuantileType) + ", Quantile must be in the range 0 to 1"
48                Case "8"

49                Case "9"

50            End Select

51            Throw ErrorMessage + LB + " to " + UB

52        End If

53        j = RoundDown((p * N) + M)
54        g = N * p + M - j

          Dim Xj As Double
          Dim Xjp1 As Double

55        Xj = Application.WorksheetFunction.Small(Sample, j)
56        If g = 0 Then
57            Xjp1 = 0
58        Else
59            Xjp1 = Application.WorksheetFunction.Small(Sample, j + 1)
60        End If

61        sGeneralisedQuantile2 = (1 - g) * Xj + g * Xjp1
62        Exit Function
ErrHandler:
63        sGeneralisedQuantile2 = "#sGeneralisedQuantile2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RoundDown
' Author     : Philip Swannell
' Date       : 13-Dec-2017
' Purpose    : Rounds a double to the largest whole number small than or equal to x.
'              Slightly tricky because "When the fractional part is exactly 0.5, CInt and CLng always round it to the nearest
'              even number. For example, 0.5 rounds to 0, and 1.5 rounds to 2"
' -----------------------------------------------------------------------------------------------------------------------
Private Function RoundDown(x As Double) As Long
1         If x = CLng(x) Then
2             RoundDown = CLng(x)
3         Else
4             RoundDown = CLng(x - 0.5)
5         End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sQuantiles
' Author    : Philip Swannell
' Date      : 06-Jun-2015
' Purpose   : Given a Sample from a population, returns estimates of the population Quantiles, i.e.
'             returns X such that Pr(x<X) = Q. Linear interpolation is used if Quantiles is
'             not equal to M/(N+1) for some positive integer M <= N, the count of Sample.
' Arguments
' Sample    : A column array of number constituting a sample from an underlying population.
' Quantiles : A single value or column array of values, must be between 1/(N+1) and N/(N+1) where N is
'             the number of elements in Sample.
' -----------------------------------------------------------------------------------------------------------------------
Function sQuantiles(Sample, Quantiles)
Attribute sQuantiles.VB_Description = "Given a Sample from a population, returns estimates of the population Quantiles, i.e. returns X such that Pr(x<X) = Q. Linear interpolation is used if Quantiles is not equal to M/(N+1) for some positive integer M <= N, the count of Sample."
Attribute sQuantiles.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim N As Long
          Dim NeedInterp() As Boolean
          Dim NS As Variant
          Dim NumConfidenceLevels As Long
          Dim Result
          Dim v As Variant
          Dim ValuesAtNs
          Dim x As Double
          Dim x1 As Double
          Dim X2 As Double
          Dim y1 As Double
          Dim y2 As Double

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti Sample, Quantiles

3         N = sNRows(Sample)

4         If sNCols(Quantiles) > 1 Then Throw "Quantiles must be a column array of numbers between 0 and 1"
5         For Each v In Quantiles
6             If Not IsNumberOrDate(v) Then Throw "Quantiles must be a column array of numbers between 0 and 1"
7             If v < 0 Or v > 1 Then Throw "Quantiles must be a column array of numbers between 0 and 1"
8             If v < 1 / (N + 1) Then Throw "Minimum allowed quantile is 1/" + CStr(N + 1)
9             If v > N / (N + 1) Then Throw "Maximum allowed quantile is " + CStr(N) + "/" + CStr(N + 1) + " = " + CStr(N / (N + 1))
10        Next v

11        NumConfidenceLevels = sNRows(Quantiles)

12        NS = sReshape(0, NumConfidenceLevels * 2, 1)
13        ReDim NeedInterp(1 To NumConfidenceLevels)

14        For i = 1 To NumConfidenceLevels
15            NeedInterp(i) = (Quantiles(i, 1) * (N + 1)) <> CLng(Quantiles(i, 1) * (N + 1))
16            If NeedInterp(i) Then
17                NS(2 * i - 1, 1) = CLng(Quantiles(i, 1) * (N + 1) - 0.5)
18                NS(2 * i, 1) = NS(2 * i - 1, 1) + 1
19            Else
20                NS(2 * i - 1, 1) = CLng(Quantiles(i, 1) * (N + 1))
21                NS(2 * i, 1) = NS(2 * i - 1, 1)
22            End If
23        Next i

24        ValuesAtNs = ThrowIfError(sNthSmallest(Sample, NS))

25        Result = sReshape(0, NumConfidenceLevels, 1)
26        For i = 1 To NumConfidenceLevels
27            If NeedInterp(i) Then
28                x1 = NS(2 * i, 1) / (N + 1)
29                X2 = NS(2 * i - 1, 1) / (N + 1)
30                x = Quantiles(i, 1)
31                y1 = ValuesAtNs(2 * i, 1)
32                y2 = ValuesAtNs(2 * i - 1, 1)
33                Result(i, 1) = y1 * (X2 - x) / (X2 - x1) + y2 * (x - x1) / (X2 - x1)
34            Else
35                Result(i, 1) = ValuesAtNs(2 * i, 1)
36            End If
37        Next i

38        sQuantiles = Result

39        Exit Function
ErrHandler:
40        sQuantiles = "#sQuantiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBasketPFE
' Author    : Philip Swannell
' Date      : 06-Jun-2015
' Purpose   : Uses moment matching to calculate Potential Future Exposure on a basket of correlated
'             log normal assets at one or more confidence levels.
' -----------------------------------------------------------------------------------------------------------------------
Function sBasketPFE(ByVal Forwards, ByVal Weights, ByVal Vols, ByVal Correlations, Time As Double, ConfidenceLevels)
          Dim CalibrationResult
          Dim i As Long
          Dim InvNormal As Double
          Dim NumConfidenceLevels As Long
          Dim Result As Variant
          Dim ThisConfidenceLevel As Double
          'Parameters of the shifted log-normal
          Dim M
          Dim S
          Dim Tau
          Dim UseNeg

1         On Error GoTo ErrHandler
2         NumConfidenceLevels = sNRows(ConfidenceLevels)
3         Result = sReshape(0, NumConfidenceLevels, 1)
4         CalibrationResult = sSLNCalibrateMinPack(Forwards, Weights, Vols, Correlations, CDbl(Time))
5         If VarType(CalibrationResult) = vbString Then
              'Calibration failure, so use Sobol!
6             sBasketPFE = sBasketPFEMC(Forwards, Weights, Vols, Correlations, Time, 127, "Sobol", ConfidenceLevels)
7             Exit Function
8         Else
9             Tau = CalibrationResult(1, 1)
10            M = CalibrationResult(2, 1)
11            S = CalibrationResult(3, 1)
12            UseNeg = CalibrationResult(4, 1)
13            For i = 1 To NumConfidenceLevels
14                ThisConfidenceLevel = IIf(UseNeg, 1 - ConfidenceLevels(i, 1), ConfidenceLevels(i, 1))
15                InvNormal = func_normsinv(CDbl(ThisConfidenceLevel))
16                Result(i, 1) = Tau + IIf(UseNeg, -1, 1) * Exp(M + S * InvNormal)
17            Next i
18        End If

19        sBasketPFE = Result
20        Exit Function
ErrHandler:
21        Throw "#sBasketPFE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
