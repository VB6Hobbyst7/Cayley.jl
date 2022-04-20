Attribute VB_Name = "modAkima"
Option Explicit

Function sSpline(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Method As String = "FMM", Optional Extrapolate As String)
Attribute sSpline.VB_Description = "Returns an array of interpolated values corresponding to the input points Xin. Values are determined by cubic spline interpolation of Xgrid and Ygrid."
Attribute sSpline.VB_ProcData.VB_Invoke_Func = " \n23"

1         On Error GoTo ErrHandler
2         Select Case LCase(Method)
              Case "akima"
3                 sSpline = AkimaSpline(Xgrid, Ygrid, Xin, Extrapolate)
4             Case "akima2"
5                 sSpline = AkimaSpline2(Xgrid, Ygrid, Xin, Extrapolate)
6             Case "fmm"
7                 sSpline = FMMSpline(Xgrid, Ygrid, Xin, Extrapolate)
8             Case "natural"
9                 sSpline = NaturalSpline(Xgrid, Ygrid, Xin, Extrapolate)
10            Case "linearquadratic", "lq"
11                sSpline = LinearQuadraticSpline(Xgrid, Ygrid, Xin, Extrapolate)
12            Case Else
13                Throw "Method not recognised. Allowed values 'Natural', 'FMM', 'Akima', 'LinearQuadratic' (or 'LQ')"
14        End Select

15        Exit Function
ErrHandler:
16        sSpline = "#sSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sPolyFit
' Author    : Philip Swannell
' Date      : 10-Mar-2020
' Purpose   : Returns the coefficients of the unique polynomial that passes through the points given by
'             the vectors Xs and Ys.
' Arguments
' Xs        : The vector of x coordinates.
' Ys        : The vector of y coordinates.
'
' Notes     : Example
'             The coefficients of the cubic polynomial that passes through the points
'             (-1,-2), (0,1), (1,10) and (2,49) is given by:
'
'             sPolyFit({-1,0,1,2},{-2,1,10,49})
'
'             which evaluates to {1;2;3;4}, representing the polynomial y = 1+2x+3x^2+4x^3
' -----------------------------------------------------------------------------------------------------------------------
Function sPolyFit(ByVal Xs, ByVal Ys)
Attribute sPolyFit.VB_Description = "Returns the coefficients of the unique polynomial that passes through the points given by the vectors Xs and Ys."
Attribute sPolyFit.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim NRx As Long, NCx As Long
          Dim NRy As Long, NCy As Long
          Dim tmp As Long
          Dim i As Long, N As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR Xs, NRx, NCx
3         If NCx > 1 Then
4             If NRx = 1 Then
5                 Xs = sArrayTranspose(Xs)
6                 tmp = NRx
7                 NRx = NCx
8                 NCx = tmp
9             End If
10        End If

11        Force2DArrayR Ys, NRy, NCy
12        If NCy > 1 Then
13            If NRy = 1 Then
14                Ys = sArrayTranspose(Ys)
15                tmp = NRy
16                NRy = NCy
17                NCy = tmp
18            End If
19        End If

20        If NRx <> NRy Or NCx <> 1 Or NCy <> 1 Then Throw "Xs and Ys must be vectors of the same length"
21        N = NRx

22        For i = 1 To N
23            If Not IsNumber(Xs(i, 1)) Then Throw "Non-number found at index " + CStr(i) + " of Xs"
24            If Not IsNumber(Ys(i, 1)) Then Throw "Non-number found at index " + CStr(i) + " of Ys"
25        Next i

26        sPolyFit = PolyFitCore(Xs, Ys)

27        Exit Function
ErrHandler:
28        sPolyFit = "#sPolyFit (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'assumes inputs have been validated. both are 2-d column vectors
Private Function PolyFitCore(Xs, Ys)
          Dim N As Long
          Dim i As Long, j As Long
          Dim a() As Double
1         On Error GoTo ErrHandler
2         N = sNRows(Xs)
3         ReDim a(1 To N, 1 To N)

4         For j = 1 To N
5             a(1, j) = 1
6         Next

7         For i = 2 To N
8             For j = 1 To N
9                 a(i, j) = a(i - 1, j) * CDbl(Xs(j, 1))
10            Next
11        Next

12        PolyFitCore = sArrayTranspose(Application.WorksheetFunction.MMult(sArrayTranspose(Ys), Application.WorksheetFunction.MInverse(a)))

13        Exit Function
ErrHandler:
14        Throw "#sPolyFitCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sPolyEval
' Author    : Philip Swannell
' Date      : 10-Mar-2020
' Purpose   : Evaluate the polynomial with coefficients Coeffs at the x-values Xs
' Arguments
' Coeffs    : The coefficients of the polynomial as a 1-row or 1-column array.
' Xs        : A number or array of numbers at which the polynomial is evaluated.
'
' Notes     : Example
'             To evaluate the polynomial y = 1+2x+3x^2 for integers from 1 to 20
'             sPolyEval(sIntegers(20),{1,2,3})
' -----------------------------------------------------------------------------------------------------------------------
Function sPolyEval(ByVal Coeffs As Variant, Xs As Variant)
Attribute sPolyEval.VB_Description = "Evaluate the polynomial with coefficients Coeffs at the x-values Xs"
Attribute sPolyEval.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim NR As Long, NC As Long, tmp As Long, i As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR Coeffs, NR, NC
3         If NC > 1 Then
4             If NR = 1 Then
5                 Coeffs = sArrayTranspose(Coeffs)
6                 tmp = NR
7                 NR = NC
8                 NC = tmp
9             Else
10                Throw "Coeffs must be one-row or 1-column array of numbers"
11            End If
12        End If

13        For i = 1 To NR
14            If Not IsNumber(Coeffs(i, 1)) Then Throw "Non-number found at index " + CStr(i) + " of Coeffs"
15        Next i

16        If VarType(Xs) < vbArray Then
17            sPolyEval = sPolyEvalCore(Coeffs, CDbl(Xs))
18        Else
              Dim NRx As Long, NCx As Long, j As Long, Res As Variant
19            Force2DArrayR Xs, NRx, NCx
20            Res = sReshape(0, NRx, NCx)
21            For i = 1 To NRx
22                For j = 1 To NCx
23                    If Not IsNumber(Xs(i, j)) Then Throw "Non-number found at index " + CStr(i) + "," + CStr(j) + " of Xs"
24                    Res(i, j) = sPolyEvalCore(Coeffs, CDbl(Xs(i, j)))
25                Next j
26            Next i
27            sPolyEval = Res
28        End If

29        Exit Function
ErrHandler:
30        sPolyEval = "#sPolyEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function sPolyEvalCore(Coeffs, x As Double)
          Dim Res As Double
          Dim i As Long
          Dim N As Long

1         On Error GoTo ErrHandler
2         N = UBound(Coeffs, 1)
3         Res = Coeffs(UBound(Coeffs, 1), 1)
4         For i = N To 2 Step -1
5             Res = Res * x + Coeffs(i - 1, 1)
6         Next

7         sPolyEvalCore = Res
8         Exit Function
ErrHandler:
9         Throw "#sPolyEvalCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CubicFit
' Author     : Philip Swannell
' Date       : 19-Mar-2020
' Purpose    : Returns the coefficients of a cubic that goes through (0,y_1), (h,y_2) and has slope s_1 at 0 and slope s_2 at h
' -----------------------------------------------------------------------------------------------------------------------
Private Function CubicFit3(H As Double, y_1 As Double, y_2 As Double, s_1 As Double, s_2 As Double)

1         On Error GoTo ErrHandler
          Dim Res(1 To 4, 1 To 1) As Double
          
2         Res(1, 1) = y_1
3         Res(2, 1) = s_1
4         Res(3, 1) = (3 * y_2 - 3 * y_1 - 2 * H * s_1 - H * s_2) / H / H
5         Res(4, 1) = (2 * y_1 - 2 * y_2 + s_1 * H + s_2 * H) / H / H / H
6         CubicFit3 = Res
7         Exit Function
ErrHandler:
8         Throw "#CubicFit3 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CubicFit
' Author     : Philip Swannell
' Date       : 06-Mar-2020
' Purpose    : Returns the coefficients of a cubic that goes through (x_1,y_1), (x_2,y_2) and has slope s_1 at x_1 and slope s_2 at x_2
' -----------------------------------------------------------------------------------------------------------------------
Private Function CubicFit(x_1 As Double, x_2 As Double, y_1 As Double, y_2 As Double, s_1 As Double, s_2 As Double)

          Dim a(1 To 4, 1 To 4) As Double
          Dim vec(1 To 4, 1 To 1) As Double

1         On Error GoTo ErrHandler
2         a(1, 1) = 1
3         a(1, 2) = x_1
4         a(1, 3) = x_1 * x_1
5         a(1, 4) = x_1 * x_1 * x_1
6         a(2, 1) = 1
7         a(2, 2) = x_2
8         a(2, 3) = x_2 * x_2
9         a(2, 4) = x_2 * x_2 * x_2
10        a(3, 1) = 0
11        a(3, 2) = 1
12        a(3, 3) = 2 * x_1
13        a(3, 4) = 3 * x_1 * x_1
14        a(4, 1) = 0
15        a(4, 2) = 1
16        a(4, 3) = 2 * x_2
17        a(4, 4) = 3 * x_2 * x_2

18        vec(1, 1) = y_1
19        vec(2, 1) = y_2
20        vec(3, 1) = s_1
21        vec(4, 1) = s_2

22        CubicFit = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(a), vec)

23        Exit Function
ErrHandler:
24        Throw "#CubicFit (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CubicFit2
' Author     : Philip Swannell
' Date       : 06-Mar-2020
' Purpose    : Returns the coefficients of a cubic that goes through (x_1,y_1), (x_2,y_2) and has second derivative ydashdash1
' and second derivative ydashdash2 at x_2
' -----------------------------------------------------------------------------------------------------------------------
Private Function CubicFit2(x_1 As Double, x_2 As Double, y_1 As Double, y_2 As Double, ydashdash1 As Double, ydashdash2 As Double)

          Dim a(1 To 4, 1 To 4) As Double
          Dim vec(1 To 4, 1 To 1) As Double

1         On Error GoTo ErrHandler
2         a(1, 1) = 1
3         a(1, 2) = x_1
4         a(1, 3) = x_1 * x_1
5         a(1, 4) = x_1 * x_1 * x_1
6         a(2, 1) = 1
7         a(2, 2) = x_2
8         a(2, 3) = x_2 * x_2
9         a(2, 4) = x_2 * x_2 * x_2
10        a(3, 1) = 0
11        a(3, 2) = 0
12        a(3, 3) = 2
13        a(3, 4) = 6 * x_1
14        a(4, 1) = 0
15        a(4, 2) = 0
16        a(4, 3) = 2
17        a(4, 4) = 6 * x_2

18        vec(1, 1) = y_1
19        vec(2, 1) = y_2
20        vec(3, 1) = ydashdash1
21        vec(4, 1) = ydashdash2

22        CubicFit2 = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(a), vec)

23        Exit Function
ErrHandler:
24        Throw "#CubicFit2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CheckSplineInputs
' Author     : Philip Swannell
' Date       : 12/03/2020
' Purpose    : Validate inputs, throwing errors if they are bad.
' Parameters :
'  Xgrid     :
'  Ygrid     :
'  Xin       :
'  N         :
'  NRXin     :
'  NCXIn     :
'  LeftOrder :
'  RightOrder:
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckSplineInputs(ByRef Xgrid, ByRef Ygrid, ByRef Xin, ByRef N As Long, ByRef NRXin As Long, ByRef NCXIn As Long, Optional LeftOrder As Long = -1, Optional RightOrder As Long = -1)
          Dim NRx As Long, NCx As Long, NRy As Long, NCy As Long, tmp As Long, i As Long, j As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR Xgrid, NRx, NCx
3         If NCx > 1 Then
4             If NRx = 1 Then
5                 Xgrid = sArrayTranspose(Xgrid)
6                 tmp = NRx
7                 NRx = NCx
8                 NCx = tmp
9             End If
10        End If

11        Force2DArrayR Ygrid, NRy, NCy
12        If NCy > 1 Then
13            If NRy = 1 Then
14                Ygrid = sArrayTranspose(Ygrid)
15                tmp = NRy
16                NRy = NCy
17                NCy = tmp
18            End If
19        End If

20        If NRx <> NRy Or NCx <> 1 Or NCy <> 1 Then Throw "Xgrid and Ygrid must be vectors of the same length"
21        N = NRx

22        For i = 1 To N
23            If Not IsNumber(Xgrid(i, 1)) Then Throw "Non-number found in Xgrid at position " + CStr(i)
24            If Not IsNumber(Ygrid(i, 1)) Then Throw "Non-number found in Ygrid at position " + CStr(i)
25        Next

26        For i = 2 To N
27            If Xgrid(i, 1) <= Xgrid(i - 1, 1) Then Throw "Xgrid must be sorted in ascending order, but element " + CStr(i) + " is not greater than element " + CStr(i - 1)
28        Next i

29        Force2DArrayR Xin, NRXin, NCXIn

30        For i = 1 To NRXin
31            For j = 1 To NCXIn
32                If Not IsNumber(Xin(i, j)) Then Throw "Non-number found in Xin at position " + CStr(i) + "," + CStr(j)
33                If LeftOrder = -1 Then
34                    If Xin(i, j) < Xgrid(1, 1) Then Throw "Left extrapolation is 'None' so all elements of Xin must be greater than or equal to the first element of Xgrid, but element " + CStr(i) + "," + CStr(j) + " is equal to " + CStr(Xin(i, j)) + " which is not greater than " + CStr(Xgrid(1, 1))
35                End If
36                If RightOrder = -1 Then
37                    If Xin(i, j) > Xgrid(N, 1) Then Throw "Right extrapolation is 'None' so all elements of Xin must be less than or equal to the last element of Xgrid, but element " + CStr(i) + "," + CStr(j) + " is equal to " + CStr(Xin(i, j)) + " which is not less than " + CStr(Xgrid(N, 1))
38                End If
39            Next
40        Next

41        Exit Function
ErrHandler:
42        Throw "#CheckSplineInputs (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AkimaSpline
' Author    : Philip Swannell
' Date      : 10-Mar-2020
' Purpose   : Fits an Akima spline to the points defined by Xgrid and Ygrid, and evaluates the spline at
'             x-values given by Xin.
' Arguments
' Xgrid     : A vector of X values of the "knots".
' Ygrid     : A vector of Y values of the "knots".
' Xin       : An array of values at which to evaluate the spline. Extrapolation is not supported.
'
' Notes     :https://en.wikipedia.org/wiki/Akima_spline
' -----------------------------------------------------------------------------------------------------------------------
Private Function AkimaSpline(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Extrapolate As String)

          Dim NRXin As Long, NCXIn As Long
          Dim tmp As Long
          Dim i As Long, j As Long, N As Long, NPrime As Long
          Dim LeftOrder As Long, RightOrder As Long

1         On Error GoTo ErrHandler

2         ParseExtrapolate Extrapolate, LeftOrder, RightOrder
3         CheckSplineInputs Xgrid, Ygrid, Xin, N, NRXin, NCXIn, LeftOrder, RightOrder

          Dim FirstXs(1 To 2, 1 To 1) As Double, LastXs(1 To 2, 1 To 1) As Double, Xs

4         tmp = Xgrid(3, 1) - Xgrid(1, 1)
5         FirstXs(1, 1) = Xgrid(1, 1) - tmp
6         FirstXs(2, 1) = Xgrid(2, 1) - tmp

7         tmp = Xgrid(N, 1) - Xgrid(N - 2, 1)
8         LastXs(1, 1) = Xgrid(N - 1, 1) + tmp
9         LastXs(2, 1) = Xgrid(N, 1) + tmp
10        Xs = sArrayStack(FirstXs, Xgrid, LastXs)

          Dim FirstYs, LastYs, Ys
          'This is quite inefficient. See Julia implementation for more succint way.
          'https://github.com/sp94/CubicSplines.jl/blob/master/src/CubicSplines.jl
11        FirstYs = sPolyEval(sPolyFit(sSubArray(Xgrid, 1, , 3), sSubArray(Ygrid, 1, , 3)), FirstXs)
12        LastYs = sPolyEval(sPolyFit(sSubArray(Xgrid, N - 2, , 3), sSubArray(Ygrid, N - 2, , 3)), LastXs)
13        Ys = sArrayStack(FirstYs, Ygrid, LastYs)

14        NPrime = N + 4

          Dim Ms() As Double, Ss() As Double
15        ReDim Ms(1 To NPrime - 1, 1 To 1)
16        ReDim Ss(1 To N, 1 To 1)

17        For i = 1 To NPrime - 1
18            Ms(i, 1) = (Ys(i + 1, 1) - Ys(i, 1)) / (Xs(i + 1, 1) - Xs(i, 1))
19        Next i

          Dim Numer As Double, Denom As Double

20        For i = 1 To N
21            Numer = Abs(Ms(i + 3, 1) - Ms(i + 2, 1)) * Ms(i + 1, 1) + _
                  Abs(Ms(i + 1, 1) - Ms(i + 0, 1)) * Ms(i + 2, 1)
22            Denom = Abs(Ms(i + 3, 1) - Ms(i + 2, 1)) + _
                  Abs(Ms(i + 1, 1) - Ms(i + 0, 1))
23            If Denom = 0 Then
24                Ss(i, 1) = (Ms(i + 1, 1) + Ms(i + 2, 1)) / 2
25            Else
26                Ss(i, 1) = Numer / Denom
27            End If
28        Next i

          Dim PolyCoeffs() As Variant
29        ReDim PolyCoeffs(0 To N)  'Array of arrays

30        For i = 1 To N - 1
31            PolyCoeffs(i) = CubicFit3(CDbl(Xgrid(i + 1, 1)) - CDbl(Xgrid(i, 1)), CDbl(Ygrid(i, 1)), CDbl(Ygrid(i + 1, 1)), Ss(i, 1), Ss(i + 1, 1))
32        Next i

33        If LeftOrder > -1 Then
34            If LeftOrder >= 2 Then
                  'Follow R's aspline function by extrapolating quadratic/cubic fitted to last three/4 knots, rather than quadratic/cubic that is the 2nd/3rd-order tangent at the end knot.
                  'Unfortunately this is not _exactly_ what the R code does, which I haven't been able to determine - since the R code calls FORTRAN, which (I discovered) is quite impenetrable.
                  'The match is exact for left extrapolation if x_2 - x_1 = x_3 - x_2
                  'and exact for right extrapolation if x_n - x_(n-1) = x_(n-1) - x_(n-2)
                  'Note that by definition of FirstXs, FirstYs etc, the first/last four (augmented) knots lie on a quadratic
35                PolyCoeffs(0) = PolyFitCore(sSubArray(Xs, , , 1 + LeftOrder), sSubArray(Ys, , , 1 + LeftOrder))
36            Else
37                PolyCoeffs(0) = PolyTangent(PolyCoeffs(1), CDbl(Xgrid(1, 1)), LeftOrder)
38            End If
39        End If
40        If RightOrder > -1 Then
41            If RightOrder >= 2 Then
                  'Follow R's aspline function...
42                PolyCoeffs(N) = PolyFitCore(sSubArray(Xs, -1 - RightOrder), sSubArray(Ys, -1 - RightOrder))
43            Else
44                PolyCoeffs(N) = PolyTangent(PolyCoeffs(N - 1), CDbl(Xgrid(N, 1)), RightOrder)
45            End If
46        End If

          Dim Result() As Double, MatchRes As Long
47        ReDim Result(1 To NRXin, 1 To NCXIn)
48        For i = 1 To NRXin
49            For j = 1 To NCXIn
50                If Xin(i, j) = Xgrid(N, 1) Then
51                    Result(i, j) = Ygrid(N, 1) ' special case
52                Else
53                    MatchRes = BinaryChop(Xgrid, CDbl(Xin(i, j)))
54                    If MatchRes = 0 Or MatchRes = N Then
55                        Result(i, j) = sPolyEvalCore(PolyCoeffs(MatchRes), CDbl(Xin(i, j)))
56                    Else
57                        Result(i, j) = sPolyEvalCore(PolyCoeffs(MatchRes), CDbl(Xin(i, j)) - CDbl(Xgrid(MatchRes, 1)))
58                    End If
59                End If
60            Next j
61        Next i

62        AkimaSpline = Result

63        Exit Function
ErrHandler:
64        Throw "#AkimaSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function QuadraticFit(x_1 As Double, x_2 As Double, x_3 As Double, y_1 As Double, y_2 As Double, y_3 As Double, x_4 As Double) As Double

          Dim X2 As Double, x3 As Double, y2 As Double, y3 As Double, x As Double

          Dim c As Double, b As Double
1         On Error GoTo ErrHandler
          'Put the origin at x_1, y_1, so we have y = bx+cx^2
2         X2 = x_2 - x_1
3         x3 = x_3 - x_1
4         x = x_4 - x_1
5         y2 = y_2 - y_1
6         y3 = y_3 - y_1

7         b = (x3 * x3 * y2 - X2 * X2 * y3) / (X2 * x3 * x3 - x3 * X2 * X2)
8         c = (x3 * y2 - X2 * y3) / (x3 * X2 * X2 - X2 * x3 * x3)

9         QuadraticFit = x * (b + c * x) + y_1

10        Exit Function
ErrHandler:
11        Throw "#QuadraticFit (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function AkimaSpline2(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Extrapolate As String)

          Dim NRXin As Long, NCXIn As Long
          Dim tmp As Long
          Dim i As Long, j As Long, N As Long, NPrime As Long
          Dim LeftOrder As Long, RightOrder As Long

1         On Error GoTo ErrHandler

2         ParseExtrapolate Extrapolate, LeftOrder, RightOrder
3         CheckSplineInputs Xgrid, Ygrid, Xin, N, NRXin, NCXIn, LeftOrder, RightOrder

          Dim FirstXs(1 To 2, 1 To 1) As Double, LastXs(1 To 2, 1 To 1) As Double, Xs
          Dim FirstYs(1 To 2, 1 To 1) As Double, LastYs(1 To 2, 1 To 1) As Double, Ys

4         tmp = Xgrid(3, 1) - Xgrid(1, 1)
5         FirstXs(1, 1) = Xgrid(1, 1) - tmp
6         FirstXs(2, 1) = Xgrid(2, 1) - tmp

7         tmp = Xgrid(N, 1) - Xgrid(N - 2, 1)
8         LastXs(1, 1) = Xgrid(N - 1, 1) + tmp
9         LastXs(2, 1) = Xgrid(N, 1) + tmp
10        Xs = sArrayStack(FirstXs, Xgrid, LastXs)

11        For i = 1 To 2
12            FirstYs(i, 1) = QuadraticFit(CDbl(Xgrid(1, 1)), CDbl(Xgrid(2, 1)), CDbl(Xgrid(3, 1)), CDbl(Ygrid(1, 1)), CDbl(Ygrid(2, 1)), CDbl(Ygrid(3, 1)), FirstXs(i, 1))
13            LastYs(i, 1) = QuadraticFit(CDbl(Xgrid(N - 2, 1)), CDbl(Xgrid(N - 1, 1)), CDbl(Xgrid(N, 1)), CDbl(Ygrid(N - 2, 1)), CDbl(Ygrid(N - 1, 1)), CDbl(Ygrid(N, 1)), LastXs(i, 1))
14        Next

15        Ys = sArrayStack(FirstYs, Ygrid, LastYs)

16        NPrime = N + 4

          Dim Ms() As Double, Ss() As Double
17        ReDim Ms(1 To NPrime - 1, 1 To 1)
18        ReDim Ss(1 To N, 1 To 1)

19        For i = 1 To NPrime - 1
20            Ms(i, 1) = (Ys(i + 1, 1) - Ys(i, 1)) / (Xs(i + 1, 1) - Xs(i, 1))
21        Next i

          Dim Numer As Double, Denom As Double

22        For i = 1 To N
23            Numer = Abs(Ms(i + 3, 1) - Ms(i + 2, 1)) * Ms(i + 1, 1) + _
                  Abs(Ms(i + 1, 1) - Ms(i + 0, 1)) * Ms(i + 2, 1)
24            Denom = Abs(Ms(i + 3, 1) - Ms(i + 2, 1)) + _
                  Abs(Ms(i + 1, 1) - Ms(i + 0, 1))
25            If Denom = 0 Then
26                Ss(i, 1) = (Ms(i + 1, 1) + Ms(i + 2, 1)) / 2
27            Else
28                Ss(i, 1) = Numer / Denom
29            End If
30        Next i

          Dim PolyCoeffs() As Variant
31        ReDim PolyCoeffs(0 To N)  'Array of arrays

32        For i = 1 To N - 1
33            PolyCoeffs(i) = CubicFit(CDbl(Xgrid(i, 1)), CDbl(Xgrid(i + 1, 1)), CDbl(Ygrid(i, 1)), CDbl(Ygrid(i + 1, 1)), Ss(i, 1), Ss(i + 1, 1))
34        Next i

35        If LeftOrder > -1 Then
36            If LeftOrder >= 2 Then
                  'Follow R's aspline function by extrapolating quadratic/cubic fitted to last three/4 knots, rather than quadratic/cubic that is the 2nd/3rd-order tangent at the end knot.
                  'Unfortunately this is not _exactly_ what the R code does, which I haven't been able to determine - since the R code calls FORTRAN, which (I discovered) is quite impenetrable.
                  'The match is exact for left extrapolation if x_2 - x_1 = x_3 - x_2
                  'and exact for right extrapolation if x_n - x_(n-1) = x_(n-1) - x_(n-2)
                  'Note that by definition of FirstXs, FirstYs etc, the first/last four (augmented) knots lie on a quadratic
37                PolyCoeffs(0) = PolyFitCore(sSubArray(Xs, , , 1 + LeftOrder), sSubArray(Ys, , , 1 + LeftOrder))
38            Else
39                PolyCoeffs(0) = PolyTangent(PolyCoeffs(1), CDbl(Xgrid(1, 1)), LeftOrder)
40            End If
41        End If
42        If RightOrder > -1 Then
43            If RightOrder >= 2 Then
                  'Follow R's aspline function...
44                PolyCoeffs(N) = PolyFitCore(sSubArray(Xs, -1 - RightOrder), sSubArray(Ys, -1 - RightOrder))
45            Else
46                PolyCoeffs(N) = PolyTangent(PolyCoeffs(N - 1), CDbl(Xgrid(N, 1)), RightOrder)
47            End If
48        End If

          Dim Result() As Double, MatchRes As Long
49        ReDim Result(1 To NRXin, 1 To NCXIn)
50        For i = 1 To NRXin
51            For j = 1 To NCXIn
52                If Xin(i, j) = Xgrid(N, 1) Then
53                    Result(i, j) = Ygrid(N, 1) ' special case
54                Else
55                    MatchRes = BinaryChop(Xgrid, CDbl(Xin(i, j)))
56                    Result(i, j) = sPolyEvalCore(PolyCoeffs(MatchRes), CDbl(Xin(i, j)))
57                End If
58            Next j
59        Next i

60        AkimaSpline2 = Result

61        Exit Function
ErrHandler:
62        Throw "#AkimaSpline2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BinaryChop
' Author     : Philip Swannell
' Date       : 09-Mar-2020
' Purpose    : returns index of greatest element of Xgrid less than or equal to y, or zero if y is less than the first value in xgrid
' Parameters :
'  Xgrid:   1 column 2-d array of numbers, ascending
'  y    :
' -----------------------------------------------------------------------------------------------------------------------
Private Function BinaryChop(Xgrid, y As Double)

          Dim Lower As Long
          Dim Middle As Long
          Dim Upper As Long

1         On Error GoTo ErrHandler
          
2         Lower = LBound(Xgrid, 1)
3         Upper = UBound(Xgrid, 1)

4         If y < Xgrid(Lower, 1) Then
5             BinaryChop = Lower - 1
6             Exit Function
7         End If

LoopStart:
8         Middle = (Upper + Lower) / 2

9         If Upper - Lower = 1 Then
10            If y >= Xgrid(Upper, 1) Then
11                BinaryChop = Upper
12            Else
13                BinaryChop = Lower
14            End If
15            Exit Function
16        End If

17        If y = Xgrid(Middle, 1) Then
18            BinaryChop = Middle
19            Exit Function
20        ElseIf y > Xgrid(Middle, 1) Then
21            Lower = Middle
22        Else
23            Upper = Middle
24        End If
25        GoTo LoopStart

26        Exit Function
ErrHandler:
27        Throw "#BinaryChop (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LinearQuadraticSpline
' Author     : Philip Swannell
' Date       : 12/03/2020
' Purpose    : "Linear Quadratic Spline" (my own name). For four knots (x1,y1) to (x4,y4) then if x2 < x_in <x3
'               result is the linearly-weighted sum of two quadratics.
'               w1 = (x3-x_in)/(x3-x2)
'               w2 = 1-w1
'               result = w1*f1(x_in) + w2 * f2(x_in)
'               where f1 is the quadratic through (x1,y1), (x2,y2), (x3,y3)
'               and   f2 is the quadratic through (x2,y2), (x3,y3), (x4,y4)
' Between the first two and last two x's the value returned is the value of the quadratic through the first (last) three knots.
' -----------------------------------------------------------------------------------------------------------------------
Private Function LinearQuadraticSpline(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Extrapolate As String = "None")
          Dim NRXin As Long, NCXIn As Long
          Dim i As Long, j As Long, N As Long
          Dim LeftOrder As Long, RightOrder As Long

1         On Error GoTo ErrHandler

2         ParseExtrapolate Extrapolate, LeftOrder, RightOrder
3         CheckSplineInputs Xgrid, Ygrid, Xin, N, NRXin, NCXIn, LeftOrder, RightOrder
4         If LeftOrder = 3 Or RightOrder = 3 Then
5             Throw "Cubic Extrapolation is not supported for LinearQuadraticSpline, since the lrftmost and rightmost intervals are quadratic rather than cubic"
6         End If

          Dim Coeffs() As Variant
7         ReDim Coeffs(1 To N - 2)
          Dim ThreeXs() As Double, ThreeYs() As Double
8         ReDim ThreeXs(1 To 3, 1 To 1)
9         ReDim ThreeYs(1 To 3, 1 To 1)

10        For i = 1 To N - 2
11            ThreeXs(1, 1) = Xgrid(i, 1)
12            ThreeXs(2, 1) = Xgrid(i + 1, 1)
13            ThreeXs(3, 1) = Xgrid(i + 2, 1)
14            ThreeYs(1, 1) = Ygrid(i, 1)
15            ThreeYs(2, 1) = Ygrid(i + 1, 1)
16            ThreeYs(3, 1) = Ygrid(i + 2, 1)
17            Coeffs(i) = PolyFitCore(ThreeXs, ThreeYs)
18        Next

19        If LeftOrder >= 0 Then
              Dim LeftCoeffs As Variant
20            LeftCoeffs = PolyTangent(Coeffs(1), CDbl(Xgrid(1, 1)), LeftOrder)
21        End If
22        If RightOrder >= 0 Then
              Dim RightCoeffs As Variant
23            RightCoeffs = PolyTangent(Coeffs(N - 2), CDbl(Xgrid(N, 1)), RightOrder)
24        End If

          Dim Result() As Double, MatchRes As Long
25        ReDim Result(1 To NRXin, 1 To NCXIn)
26        For i = 1 To NRXin
27            For j = 1 To NCXIn
28                MatchRes = BinaryChop(Xgrid, CDbl(Xin(i, j)))
29                Select Case MatchRes
                      Case 0
                          'doing left extrapolation
30                        Result(i, j) = sPolyEvalCore(LeftCoeffs, CDbl(Xin(i, j)))
31                    Case 1
                          'Only have one quadratic available - the "leftmost"
32                        Result(i, j) = sPolyEvalCore(Coeffs(1), CDbl(Xin(i, j)))
33                    Case N - 1
                          'Only have one quadratic available - the "rightmost"
34                        Result(i, j) = sPolyEvalCore(Coeffs(N - 2), CDbl(Xin(i, j))) ' only have one quadratic available to fit to
35                    Case N
                          'doing right extrapolation
36                        Result(i, j) = sPolyEvalCore(RightCoeffs, CDbl(Xin(i, j)))
37                    Case Else
                          'Standard case. Linear weight of two quadratic fits.
                          Dim Weight1 As Double, Weight2 As Double
38                        Weight1 = (Xgrid(MatchRes + 1, 1) - Xin(i, j)) / (Xgrid(MatchRes + 1, 1) - Xgrid(MatchRes, 1))
39                        Weight2 = (Xin(i, j) - Xgrid(MatchRes, 1)) / (Xgrid(MatchRes + 1, 1) - Xgrid(MatchRes, 1))
40                        Result(i, j) = Weight1 * sPolyEvalCore(Coeffs(MatchRes - 1), CDbl(Xin(i, j))) + _
                              Weight2 * sPolyEvalCore(Coeffs(MatchRes), CDbl(Xin(i, j)))
41                End Select
42            Next j
43        Next i

44        LinearQuadraticSpline = Result

45        Exit Function
ErrHandler:
46        Throw "#LinearQuadraticSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'Natural Spline implemented using approach described at "Cubic Spline Part B"
'http://www.vbforums.com/showthread.php?480806-Cubic-Spline-Tutorial
Private Function NaturalSpline(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Extrapolate As String = "None")
          Dim NRXin As Long, NCXIn As Long
          Dim i As Long, j As Long, N As Long
          Dim a() As Double
          Dim Hs() As Double
          Dim LeftOrder As Long
          Dim RightOrder As Long

1         On Error GoTo ErrHandler

2         ParseExtrapolate Extrapolate, LeftOrder, RightOrder
3         CheckSplineInputs Xgrid, Ygrid, Xin, N, NRXin, NCXIn, LeftOrder, RightOrder

          Dim Coeffs() As Variant
4         ReDim Coeffs(1 To N - 2)

5         ReDim a(1 To N, 1 To N)
6         ReDim Hs(1 To N - 1)

7         For i = 1 To N - 1
8             Hs(i) = Xgrid(i + 1, 1) - Xgrid(i, 1)
9         Next

10        a(1, 1) = 1
11        For i = 2 To N - 1
12            a(i, i - 1) = Hs(i - 1)
13            a(i, i) = 2 * (Hs(i - 1) + Hs(i))
14            a(i, i + 1) = Hs(i)
15        Next
16        a(N, N) = 1

          Dim b() As Double
17        ReDim b(1 To N, 1 To 1)
18        b(1, 1) = 0
19        For i = 2 To N - 1
20            b(i, 1) = 6 * ((Ygrid(i + 1, 1) - Ygrid(i, 1)) / Hs(i) - (Ygrid(i, 1) - Ygrid(i - 1, 1)) / Hs(i - 1))
21        Next
22        b(N, 1) = 0

          Dim yDoubleDash As Variant
          Dim AInverse

23        AInverse = Application.WorksheetFunction.MInverse(a)
24        yDoubleDash = Application.WorksheetFunction.MMult(AInverse, b)

          Dim PolyCoeffs() As Variant
25        ReDim PolyCoeffs(0 To N)  'Array of arrays

26        For i = 1 To N - 1
27            PolyCoeffs(i) = CubicFit2(CDbl(Xgrid(i, 1)), CDbl(Xgrid(i + 1, 1)), CDbl(Ygrid(i, 1)), CDbl(Ygrid(i + 1, 1)), CDbl(yDoubleDash(i, 1)), CDbl(yDoubleDash(i + 1, 1)))
28        Next i

29        If LeftOrder > -1 Then
30            PolyCoeffs(0) = PolyTangent(PolyCoeffs(1), CDbl(Xgrid(1, 1)), LeftOrder)
31        End If
32        If RightOrder > -1 Then
33            PolyCoeffs(N) = PolyTangent(PolyCoeffs(N - 1), CDbl(Xgrid(N, 1)), RightOrder)
34        End If

          Dim Result() As Double, MatchRes As Long
35        ReDim Result(1 To NRXin, 1 To NCXIn)
36        For i = 1 To NRXin
37            For j = 1 To NCXIn
38                If Xin(i, j) = Xgrid(N, 1) Then
39                    Result(i, j) = Ygrid(N, 1) ' special case
40                Else
41                    MatchRes = BinaryChop(Xgrid, CDbl(Xin(i, j)))
42                    Result(i, j) = sPolyEvalCore(PolyCoeffs(MatchRes), CDbl(Xin(i, j)))
43                End If
44            Next j
45        Next i

46        NaturalSpline = Result

47        Exit Function
ErrHandler:
48        Throw "#NaturalSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ParseExtrapolate(Extrap As String, ByRef LeftOrder As Long, RightOrder As Long)
1         On Error GoTo ErrHandler
2         If InStr(Extrap, ",") = 0 Then
3             LeftOrder = ExtrapToOrder(Extrap)
4             RightOrder = LeftOrder
5         Else
6             LeftOrder = ExtrapToOrder(sStringBetweenStrings(Extrap, , ","))
7             RightOrder = ExtrapToOrder(sStringBetweenStrings(Extrap, ","))
8         End If

9         Exit Function
ErrHandler:
10        Throw "#ParseExtrapolate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ExtrapToOrder(Extrap As String)
1         On Error GoTo ErrHandler
2         Select Case LCase(Extrap)

              Case "none", "", "-1"
3                 ExtrapToOrder = -1
4             Case "flat", "0"
5                 ExtrapToOrder = 0
6             Case "linear", "1"
7                 ExtrapToOrder = 1
8             Case "quadratic", "2"
9                 ExtrapToOrder = 2
10            Case "cubic", "3"
11                ExtrapToOrder = 3
12            Case Else
13                Throw "Extrapolation string '" + Extrap + "' not recognised. Allowed values: 'None', 'Flat', 'Linear', 'Quadratic', 'Cubic', or comma delimited pair for different left and right extrapolation"
14        End Select

15        Exit Function
ErrHandler:
16        Throw "#ExtrapToOrder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'FMM = Forsythe, G. E., Malcolm, M. A. and Moler, C. B. (1977). Computer Methods for Mathematical Computations. Wiley.
'Piecewise cubic with continuous first and sexond derivatives at the knots (as per natural spline). Rather than
'setting the second derivative to zero at the end points this method sets the third derivative to that of the cubic defined by the first (last) four points.
Private Function FMMSpline(ByVal Xgrid, ByVal Ygrid, ByVal Xin, Optional Extrapolate As String = "None")
          Dim NRXin As Long, NCXIn As Long
          Dim i As Long, j As Long, N As Long
          Dim a() As Double
          Dim Hs() As Double
          Dim LeftOrder As Long
          Dim RightOrder As Long

1         On Error GoTo ErrHandler

2         ParseExtrapolate Extrapolate, LeftOrder, RightOrder
3         CheckSplineInputs Xgrid, Ygrid, Xin, N, NRXin, NCXIn, LeftOrder, RightOrder

          Dim Coeffs() As Variant
4         ReDim Coeffs(1 To N - 2)

5         ReDim a(1 To N, 1 To N)
6         ReDim Hs(1 To N - 1)

7         For i = 1 To N - 1
8             Hs(i) = Xgrid(i + 1, 1) - Xgrid(i, 1)
9         Next

          'Here we amend the matrix formula used for natural spline. For natural spline set the second derivative at x1 to zero
          'but for FMM we match third derivative at the end points, so if we know the third derivative (constant for cubics) at x1
          'then we know the difference between the second derivative at x1 and the second derivative at x2 - it's simply the third
          'derivative at x1 multiplied by h1 = x2-x1

10        a(1, 1) = -1
11        a(1, 2) = 1
12        For i = 2 To N - 1
13            a(i, i - 1) = Hs(i - 1)
14            a(i, i) = 2 * (Hs(i - 1) + Hs(i))
15            a(i, i + 1) = Hs(i)
16        Next
17        a(N, N - 1) = -1
18        a(N, N) = 1

          'Need the 3rd derivative (a constant) of the spline that fits the first 4 and last 4 points
          Dim yddd1 As Double, ydddlast As Double

19        yddd1 = sPolyFit(sSubArray(Xgrid, 1, 1, 4), sSubArray(Ygrid, 1, 1, 4))(4, 1) * 6
20        ydddlast = sPolyFit(sSubArray(Xgrid, -4, 1, 4), sSubArray(Ygrid, -4, 1, 4))(4, 1) * 6

          Dim b() As Double
21        ReDim b(1 To N, 1 To 1)
22        b(1, 1) = yddd1 * Hs(1) 'see explanation above
23        For i = 2 To N - 1
24            b(i, 1) = 6 * ((Ygrid(i + 1, 1) - Ygrid(i, 1)) / Hs(i) - (Ygrid(i, 1) - Ygrid(i - 1, 1)) / Hs(i - 1))
25        Next
26        b(N, 1) = ydddlast * Hs(N - 1)

          Dim yDoubleDash As Variant
          Dim AInverse

27        AInverse = Application.WorksheetFunction.MInverse(a)
28        yDoubleDash = Application.WorksheetFunction.MMult(AInverse, b)

          Dim PolyCoeffs() As Variant
29        ReDim PolyCoeffs(0 To N)  'Array of arrays

30        For i = 1 To N - 1
31            PolyCoeffs(i) = CubicFit2(CDbl(Xgrid(i, 1)), CDbl(Xgrid(i + 1, 1)), CDbl(Ygrid(i, 1)), CDbl(Ygrid(i + 1, 1)), CDbl(yDoubleDash(i, 1)), CDbl(yDoubleDash(i + 1, 1)))
32        Next i

33        If LeftOrder > -1 Then
34            PolyCoeffs(0) = PolyTangent(PolyCoeffs(1), CDbl(Xgrid(1, 1)), LeftOrder)
35        End If
36        If RightOrder > -1 Then
37            PolyCoeffs(N) = PolyTangent(PolyCoeffs(N - 1), CDbl(Xgrid(N, 1)), RightOrder)
38        End If

          Dim Result() As Double, MatchRes As Long
39        ReDim Result(1 To NRXin, 1 To NCXIn)
40        For i = 1 To NRXin
41            For j = 1 To NCXIn
42                If Xin(i, j) = Xgrid(N, 1) Then
43                    Result(i, j) = Ygrid(N, 1) ' special case
44                Else
45                    MatchRes = BinaryChop(Xgrid, CDbl(Xin(i, j)))
46                    Result(i, j) = sPolyEvalCore(PolyCoeffs(MatchRes), CDbl(Xin(i, j)))
47                End If
48            Next j
49        Next i

50        FMMSpline = Result

51        Exit Function
ErrHandler:
52        Throw "#FMMSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PolyTangent
' Author     : Philip Swannell
' Date       : 12/03/2020
' Purpose    : Returns the tangent polynomial.
' Parameters :
'  InputPoly:
'  x        : x value at which the tangent polynomial "touches" the input polynomial
'  order    : the order of the tanget. 1 = linear 2 = quadratic etc
' -----------------------------------------------------------------------------------------------------------------------
Private Function PolyTangent(InputPoly, x As Double, order As Long)
          Dim N As Long, i As Long, j  As Long, k As Long, M As Long
          Dim DerivsAtX() As Double
          Dim a() As Double
          Dim AInv
          Dim Res

1         On Error GoTo ErrHandler
2         Force2DArrayR InputPoly, M

3         N = order + 1
4         ReDim a(1 To N, 1 To N)

5         For i = 1 To N
6             For j = i To N
7                 a(i, j) = x ^ (j - i)
8                 If i > 1 Then
9                     For k = (j - 1) To (j - 1 - i + 2) Step -1
10                        a(i, j) = a(i, j) * k
11                    Next k
12                End If
13            Next j
14        Next i

15        ReDim DerivsAtX(1 To N, 1 To 1)
16        For i = 1 To N
17            DerivsAtX(i, 1) = sPolyEvalCore(PolyDerivCore(InputPoly, i - 1), x)
18        Next i

19        AInv = Application.WorksheetFunction.MInverse(a)

20        Res = Application.WorksheetFunction.MMult(AInv, DerivsAtX)
21        If order = 0 Then
22            Force2DArray Res
23        End If
24        PolyTangent = Res
25        Exit Function
ErrHandler:
26        Throw "#PolyTangent (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PolyDerivCore
' Author     : Philip Swannell
' Date       : 12/03/2020
' Purpose    : Differentiate a polynomial
' Parameters :
'  InputPoly:
'  order    :
' -----------------------------------------------------------------------------------------------------------------------
Private Function PolyDerivCore(ByVal InputPoly, order As Long)
          Dim i As Long, N As Long, k As Long

1         On Error GoTo ErrHandler

2         N = sNRows(InputPoly)

3         For i = N To order + 1 Step -1
4             For k = i - 1 To i - order Step -1
5                 InputPoly(i, 1) = InputPoly(i, 1) * k
6             Next k
7         Next i
8         PolyDerivCore = sSubArray(InputPoly, order + 1)

9         Exit Function
ErrHandler:
10        Throw "#PolyDerivCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function sPolyDeriv(ByVal Coeffs, Optional order As Long = 1)
1         On Error GoTo ErrHandler
2         Force2DArrayR Coeffs
3         If sNCols(Coeffs) > 1 Then
4             If sNRows(Coeffs) = 1 Then
5                 Coeffs = sArrayTranspose(Coeffs)
6             Else
7                 Throw "Coeffs must have either 1 row or 1 column"
8             End If
9         End If
10        sPolyDeriv = PolyDerivCore(Coeffs, order)
11        Exit Function
ErrHandler:
12        sPolyDeriv = "#sPolyDeriv (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

