Attribute VB_Name = "modBarriers"
'Standard Barrier Options with rebates. Black-Scholes valuation
'See "The Complete Guide to Option Pricing Formulas, Second Edition, Espen Garder Haug" Page 152

Option Explicit

Private Enum EnmBarrierOptionStyle
    bos_CallDownIn = 1
    bos_CallUpIn = 2
    bos_PutDownIn = 3
    bos_PutUpIn = 4
    bos_CallDownOut = 5
    bos_CallUpOut = 6
    bos_PutDownOut = 7
    bos_PutUpOut = 8
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBarrierOption
' Author    : Philip Swannell
' Date      : 11-Feb-2021
' Purpose   : Black-Scholes valuation of a standard barrier option, with rebate. All arguments may be
'             arrays.
' Arguments
' OptionStyle: The "style" of the option. Allowed values: 'CallDownIn', 'CallDownOut', 'CallUpIn',
'             'CallUpOut', 'PutDownIn', 'PutDownOut', 'PutUpIn', 'PutUpOut`. Case
'             insensitive. Also allow three-letter abbreviations such as 'Cdi`.
' Spot      : The current spot level.
' Strike    : The option strike.
' Barrier   : The barrier level. If the underlying hits this level then the option is "switched on" (in
'             options) or "ceases to exist" (out options).
' Rebate    : Rebate level. For out options, the rebate is paid to the option holder immediately when
'             the barrier is hit. For in options, the holder is paid at maturity if the
'             barrier is not hit.
' Vol       : The annualised volatility.
' Time      : The time in years to maturity. If time is negative then the function returns zero.
' DF        : The discount factor at maturity.
' DivYield  : The dividend yield of the underlying. The model assumes this dividend is paid
'             continuously.
' AlreadyHit: A Boolean. If true then out options return a zero value, since rebate is assumed to be "in
'             the past", and in options return the Black-Scholes value of a European
'             option.
'
' Notes     : When the barrier has not yet been hit then valuation is via closed form solution as given
'             in "The Complete Guide to Option Pricing Formulas, 2nd Edition" Haug, 2007.
'             p152.
'
'             When input Spot and Barrier imply that the barrier has already been hit then
'             the returned value is Rebate (for out options) or the value of a European
'             option (for in options). But see also the "AlreadyHit" argument which if true
'             yields a return of _zero_ (for out options) or the value of a European for in
'             options.
' -----------------------------------------------------------------------------------------------------------------------
Function sBarrierOption(OptionStyle, Spot, Strike, Barrier, Rebate, Vol, Time, DF, DivYield, Optional AlreadyHit As Boolean = False) As Variant
Attribute sBarrierOption.VB_Description = "Black-Scholes valuation of a standard barrier option, with rebate. All arguments may be arrays."
Attribute sBarrierOption.VB_ProcData.VB_Invoke_Func = " \n29"
1         On Error GoTo ErrHandler

2         If VarType(OptionStyle) < vbArray And VarType(Spot) < vbArray And VarType(Strike) < vbArray And VarType(Barrier) < vbArray _
              And VarType(Rebate) < vbArray And VarType(Vol) < vbArray And VarType(Time) < vbArray And VarType(DF) < vbArray And VarType(DivYield) < vbArray And VarType(AlreadyHit) < vbArray Then
3             sBarrierOption = BarrierOption(OptionStyle, Spot, Strike, Barrier, Rebate, Vol, Time, DF, DivYield, AlreadyHit)
4         Else
              'DF and Time having different sizes is very likely to be an error.
5             If sNRows(DF) <> sNRows(Time) Or sNCols(DF) <> sNCols(Time) Then Throw "DF and Time arguments should have the same dimensions"
6             sBarrierOption = ThrowIfError(Broadcast(FuncIdBarrierOption, OptionStyle, Spot, Strike, Barrier, Rebate, Vol, Time, DF, DivYield, AlreadyHit))
7         End If
8         Exit Function
ErrHandler:
9         sBarrierOption = "#sBarrierOption (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'The "core function" wrapped by sBarrierOption.
Function BarrierOption(ByVal OptionStyle As Variant, Spot, Strike, Barrier, Rebate, Vol, Time, DF, DivYield, AlreadyHit)
          
          Dim S As Double 'Spot
          Dim x As Double 'Strike
          Dim H As Double 'Barrier
          Dim k As Double 'Rebate
          Dim sigma As Double 'Vol
          Dim s2 As Double  'vol squared
          Dim srt As Double 'vol * root(t)
          Dim t As Double 'Time
          Dim R As Double 'Interest rate
          Dim b As Double 'Cost of carry = r - Dividend yield
          
          Dim eta As Double
          Dim phi As Double
          Dim x1 As Double
          Dim X2 As Double
          Dim y1 As Double
          Dim y2 As Double
          Dim mu As Double
          Dim lambda As Double
          Dim lambda2 As Double
          Dim z As Double
          
          Dim a As Double
          Dim B_ As Double 'Underscore appended since there is also a variable b. VBA code is not case sensitive...
          Dim c As Double
          Dim D As Double
          Dim E As Double
          Dim F As Double
          
          Dim Value As Double

1         On Error GoTo ErrHandler
          
          'Validate inputs. Gives friendly error messages, including element-wise error messages when passing arrays into sBarrierOption
2         OptionStyle = ParseOptionStyle(CStr(OptionStyle))
3         CheckIsNumber Spot, "Spot"
4         CheckIsNumber Strike, "Strike"
5         CheckIsNumber Barrier, "Barrier"
6         CheckIsNumber Rebate, "Rebate"
7         CheckIsNumber Vol, "Vol"
8         CheckIsNumber Time, "Time"
9         CheckIsNumber DF, "DF"
10        CheckIsNumber DivYield, "DivYield"
11        CheckIsBool AlreadyHit, "AlreadyHit"

12        If Strike <= 0 Then Throw "Strike must be positive"
13        If Barrier <= 0 Then Throw "Barrier must be positive"
14        If Spot <= 0 Then Throw "Spot must be positive"
15        If Time < 0 Then
16            Value = 0
17            GoTo EarlyExit
18        End If

          'Match notation on p152 of "The Complete Guide to Option Pricing Formulas"
          'http://1.droppdf.com/files/RvBRE/the-complete-guide-to-option-pricing-formulas-2007.pdf
19        S = Spot
20        x = Strike
21        H = Barrier
22        k = Rebate
23        t = Time
          'Log is natural log
24        R = -Log(DF) / t
25        b = R - DivYield
26        sigma = Vol
          'Sqr is square root
27        srt = Vol * Sqr(t)
28        s2 = Vol * Vol

29        Select Case OptionStyle
              Case bos_CallDownIn, bos_CallDownOut
30                eta = 1
31                phi = 1
32            Case bos_CallUpIn, bos_CallUpOut
33                eta = -1
34                phi = 1
35            Case bos_PutDownIn, bos_PutDownOut
36                eta = 1
37                phi = -1
38            Case bos_PutUpIn, bos_PutUpOut
39                eta = -1
40                phi = -1
41        End Select
          
42        If AlreadyHit Then
43            Select Case OptionStyle
                  Case bos_CallDownOut, bos_CallUpOut, bos_PutDownOut, bos_PutUpOut
44                    BarrierOption = 0 'i.e. Rebate has been paid in the past, so not included in value in this case.
45                Case bos_CallDownIn, bos_CallUpIn
46                    BarrierOption = BSOpt(1, S * Exp((R - b) * t), x, sigma, t) * Exp(-R * t)
47                Case bos_PutDownIn, bos_PutUpIn
48                    BarrierOption = BSOpt(-1, S * Exp((R - b) * t), x, sigma, t) * Exp(-R * t)
49            End Select
50            Exit Function
51        End If

          'When barrier has been triggered to vanilla option
52        If ((S < H And OptionStyle = bos_CallDownIn)) Or _
              (S > H And OptionStyle = bos_CallUpIn) Then
53            Value = BSOpt(1, S * Exp((R - b) * t), x, sigma, t) * Exp(-R * t)
54            GoTo EarlyExit
55        ElseIf ((S < H And OptionStyle = bos_PutDownIn)) Or _
              (S > H And OptionStyle = bos_PutUpIn) Then
56            Value = BSOpt(-1, S * Exp((R - b) * t), x, sigma, t) * Exp(-R * t)
57            GoTo EarlyExit
              'When barrier triggered into immediate payment
58        ElseIf (S < H And OptionStyle = bos_CallDownOut) Or _
              (S > H And OptionStyle = bos_CallUpOut) Or _
              S < H And OptionStyle = bos_PutDownOut Or _
              S > H And OptionStyle = bos_PutUpOut Then
59            Value = k
60            GoTo EarlyExit
61        End If

62        mu = (b - s2 / 2) / s2
          
          'F is needed if and only if the trade is an out option with non-zero rebate
          'and lambda is only needed to calculate F. Lambda is the square root of a
          'quantity that may be negative if rates are negative, so we take care.
          Dim LambdaNeeded As Boolean
63        If k <> 0 Then
64            Select Case OptionStyle
                  Case bos_CallDownOut, bos_CallUpOut, bos_PutDownOut, bos_PutUpOut
65                    LambdaNeeded = True
66            End Select
67        End If

68        If LambdaNeeded Then
69            lambda2 = mu * mu + (2 * R / s2)
70            If lambda2 < 0 Then
71                Throw "Error calculating lambda. Cannot take square root of negative value"
72            Else
73                lambda = Sqr(lambda2)
74            End If
75        End If

          'Code in this section quite inefficient, as we are calculating quantities that may not be needed.
76        x1 = Log(S / x) / srt + (1 + mu) * srt
77        X2 = Log(S / H) / srt + (1 + mu) * srt
78        y1 = Log(H * H / (S * x)) / srt + (1 + mu) * srt
79        y2 = Log(H / S) / srt + (1 + mu) * srt
80        z = (Log(H / S) / srt) + lambda * srt

81        a = phi * S * Exp((b - R) * t) * N(phi * x1) - phi * x * Exp(-R * t) * N(phi * x1 - phi * srt)
82        B_ = phi * S * Exp((b - R) * t) * N(phi * X2) - phi * x * Exp(-R * t) * N(phi * X2 - phi * srt)
83        c = phi * S * Exp((b - R) * t) * (H / S) ^ (2 * (mu + 1)) * N(eta * y1) - phi * x * Exp(-R * t) * (H / S) ^ (2 * mu) * N(eta * y1 - eta * srt)
84        D = phi * S * Exp((b - R) * t) * (H / S) ^ (2 * (mu + 1)) * N(eta * y2) - phi * x * Exp(-R * t) * (H / S) ^ (2 * mu) * N(eta * y2 - eta * srt)

85        If k <> 0 Then
86            E = k * Exp(-R * t) * (N(eta * X2 - eta * srt) - (H / S) ^ (2 * mu) * N(eta * y2 - eta * srt))
87            F = k * ((H / S) ^ (mu + lambda) * N(eta * z) + (H / S) ^ (mu - lambda) * N(eta * z - 2 * eta * lambda * srt))
88        Else
89            E = 0: F = 0 'This line not really necessary. In VBA variables Dim'd as Double are initialised to zero
90        End If

91        If x > H Then
92            Select Case OptionStyle
                  Case bos_CallDownIn
93                    Value = c + E
94                Case bos_CallUpIn
95                    Value = a + E
96                Case bos_PutDownIn
97                    Value = B_ - c + D + E
98                Case bos_PutUpIn
99                    Value = a - B_ + D + E
100               Case bos_CallDownOut
101                   Value = a - c + F
102               Case bos_CallUpOut
103                   Value = F
104               Case bos_PutDownOut
105                   Value = a - B_ + c - D + F
106               Case bos_PutUpOut
107                   Value = B_ - D + F
108           End Select
109       Else
110           Select Case OptionStyle
                  Case bos_CallDownIn
111                   Value = a - B_ + D + E
112               Case bos_CallUpIn
113                   Value = B_ - c + D + E
114               Case bos_PutDownIn
115                   Value = a + E
116               Case bos_PutUpIn
117                   Value = c + E
118               Case bos_CallDownOut
119                   Value = B_ - D + F
120               Case bos_CallUpOut
121                   Value = a - B_ + c - D + F
122               Case bos_PutDownOut
123                   Value = F
124               Case bos_PutUpOut
125                   Value = a - c + F
126           End Select
127       End If

EarlyExit:
128       BarrierOption = Value

129       Exit Function
ErrHandler:
130       BarrierOption = "#BarrierOption (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function CheckIsNumber(Arg As Variant, ArgName As String)
1         Select Case VarType(Arg)
              Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
2             Case Else
3                 Throw ArgName + " must be a number"
4         End Select
End Function

Private Function CheckIsBool(Arg As Variant, ArgName As String)
1         If Not VarType(Arg) = vbBoolean Then Throw ArgName + " must be a Boolean"
End Function

'Copy of Throw, to make this module (nearly) self-contained.
Private Sub Throw(ByVal ErrorString As String)
1         Err.Raise vbObjectError + 1, , ErrorString
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BSOpt
' Author     : Philip Swannell
' Date       : 10-Feb-2021
' Purpose    : UNDISCOUNTED Black Scholes value of vanilla option. CP = 1 for Call, -1 for Put
' -----------------------------------------------------------------------------------------------------------------------
Private Function BSOpt(CP As Integer, Forward As Double, Strike As Double, Vol As Double, Time As Double) As Double
          Dim d1 As Double, d2 As Double, srt As Double
1         srt = Vol * Sqr(Time)
2         d1 = Log(Forward / Strike) / srt + srt / 2
3         d2 = d1 - srt
4         BSOpt = CP * (Forward * N(CP * d1) - Strike * N(CP * d2))
End Function

'Wrap cumulative normal distribution function from modNormal.
'Could use Application.WorksheetFunction.NormSDist, but this is equally accurate and faster.
Private Function N(z As Double) As Double
1         N = func_normsdistdd(z)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseOptionStyle
' Author     : Philip Swannell
' Date       : 10-Feb-2021
' Purpose    : Convert human-friendly string (e.g. "Down and Out Call" or "CallDownOut" or "Cdo") into an EnmBarrierOptionStyle
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseOptionStyle(OptionStyle As String)
1         On Error GoTo ErrHandler

          Const ErrMsg = "OptionStyle not recognised. Allowed = 'CallDownIn', 'CallDownOut', 'CallUpIn', 'CallUpOut', 'PutDownIn', 'PutDownOut', 'PutUpIn', 'PutUpOut' (case insensitive, spaces ignored, also three-letter abbreviations such as Cdi)."

2         OptionStyle = LCase(OptionStyle)
3         If Len(OptionStyle) > 3 Then
4             OptionStyle = Replace(OptionStyle, " ", "")
5             OptionStyle = Replace(OptionStyle, "-", "")
6             OptionStyle = Replace(OptionStyle, "and", "")
7             OptionStyle = Replace(OptionStyle, "call", "c")
8             OptionStyle = Replace(OptionStyle, "put", "p")
9             OptionStyle = Replace(OptionStyle, "out", "o")
10            OptionStyle = Replace(OptionStyle, "in", "i")
11            OptionStyle = Replace(OptionStyle, "down", "d")
12            OptionStyle = Replace(OptionStyle, "up", "u")
13        End If
14        Select Case OptionStyle
              Case "dic", "cdi"
15                OptionStyle = bos_CallDownIn
16            Case "uic", "cui"
17                OptionStyle = bos_CallUpIn
18            Case "dip", "pdi"
19                OptionStyle = bos_PutDownIn
20            Case "uip", "pui"
21                OptionStyle = bos_PutUpIn
22            Case "doc", "cdo"
23                OptionStyle = bos_CallDownOut
24            Case "uoc", "cuo"
25                OptionStyle = bos_CallUpOut
26            Case "dop", "pdo"
27                OptionStyle = bos_PutDownOut
28            Case "uop", "puo"
29                OptionStyle = bos_PutUpOut
30            Case Else
31                Throw ErrMsg
32        End Select
33        ParseOptionStyle = OptionStyle
34        Exit Function
ErrHandler:
35        Throw "#ParseOptionStyle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'PGS implementation of Hull
'John Hull, Options, Futures and Other Derivative Securities, Sixth Edition Page 533
'Results are the same as Haug, though code is simpler as handle only Down and in calls with no rebate.
Private Function CallDownInHull(Spot As Double, Strike As Double, _
        Barrier As Double, Vol As Double, _
        Time As Double, DF As Double, DivYield As Double) As Variant

          Dim S As Double
          Dim t As Double
          Dim q As Double
          Dim H As Double
          Dim R As Double
          Dim lambda As Double
          Dim y As Double
          Dim srt As Double
          Dim s2 As Double
          Dim k As Double
          Dim Cdo As Double
          Dim x1 As Double
          Dim y1 As Double
          Dim CallValue As Double
          Dim CallDownInValue As Double

1         S = Spot
2         t = Time
3         R = -Log(DF) / t
4         q = DivYield
5         srt = Vol * Sqr(t)
6         s2 = Vol * Vol
7         H = Barrier
8         k = Strike

9         If S < H Then
10            CallDownInHull = BSOpt(1, Spot * Exp((R - q) * t), k, Vol, Time) * Exp(-R * t)
11        Else
12            lambda = (R - q + s2 / 2) / s2
13            y = Log(H * H / (S * k)) / srt + lambda * srt
14            If H < k Then
15                CallDownInValue = _
                      S * Exp(-q * t) * (H / S) ^ (2 * lambda) * N(y) - _
                      k * Exp(-R * t) * (H / S) ^ (2 * lambda - 2) * N(y - srt)
16            Else
17                x1 = Log(S / H) / srt + lambda * srt
18                y1 = Log(H / S) / srt + lambda * srt

19                Cdo = _
                      S * N(x1) * Exp(-q * t) - _
                      k * Exp(-R * t) * N(x1 - srt) - _
                      S * Exp(-q * t) * (H / S) ^ (2 * lambda) * N(y1) + _
                      k * Exp(-R * t) * (H / S) ^ (2 * lambda - 2) * N(y1 - srt)
20                CallValue = BSOpt(1, S * Exp((R - q) * t), k, Vol, t) * Exp(-R * t)
21                CallDownInValue = CallValue - Cdo
22            End If
23        End If

24        CallDownInHull = CallDownInValue

End Function
