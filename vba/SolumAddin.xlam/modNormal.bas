Attribute VB_Name = "modNormal"
Option Private Module
' -----------------------------------------------------------------------------------------------------------------------
' Module    : ModNormal
' Author    : Philip Swannell
' Date      : 04-May-2015
' Purpose   : Code downloaded from http://www.spreadsheetadvice.com/2011/04/normal_distribution/
'             in workbook Normal_distribution_v3.xls. That workbook has two code modules Dis01_Normal
'             and Dis02_Lognormal. I have amalgamated the two into modNormal
' Changes made by PGS: Removed a number of variables and constants that MZTools revealed to be unused.
'                      Replaced line GoTo 300 by GoTo EndOfFunction since otherwise adding line numbers to
'                      this code via MZTools will break the code.
'                      In function func_normsdistdd, made the array A be static - achieves a small (27%) speedup.
' -----------------------------------------------------------------------------------------------------------------------
'****************************************************************************
'* Module contains:
'*  func_normsdist(z As Double)                                  standard normal cumulative distribution function
'*  func_normsdistdd(z As Double)                                more precise, but slower version
'*  func_normsdistdd2(z As Double)                               same precision, but a bit slower -- obsolete
'*  func_normsinv(p As Double)                                   inverse
'*  func_normsinvdd(p As Double)                                 more precise, but slower version
'*  func_normsdense(z as double)                                 density
'*  func_binormcum(z1 As Double, z2 As Double, rho As Double)    bivariate normal cumulative distribution function
'*  func_binormcumdd(z1 As Double, z2 As Double, rho As Double)  more precise, but slower version
'*  func_binormdense(z1 As Double, z2 As Double, rho As Double)  bivariate normal density function
'*  func_logdist(z As Double, mean As Double, stev As Double)    lognormal cumulative distribution function
'*  func_logdistdd(z As Double, mean As Double, stev As Double)  more precise, but slower version
'*  func_loginv(p As Double, mean As Double, stev As Double)     inverse
'*  func_loginvdd(p As Double, mean As Double, stev As Double)   more precise, but slower version
'*  func_logdense(z As Double, mean As Double, stev As Double)   density
'****************************************************************************
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestCumulativeNormalFunctions
' Author    : Philip Swannell
' Date      : 04-May-2015
' Purpose   : Test Cumulative distribution functions for speed
'             func_normsdist is 6 or 7 times faster than Application.WorksheetFunction.Norm_S_Dist
'             but the higher precision versions are slightly slower than Application.WorksheetFunction.Norm_S_Dist

'Tested again 20 Jan 2021 (after use-a-static speedup to func_normsdistdd) and got:

'Application.WorksheetFunction.Norm_S_Dist: time for 120,000 calls:     0.299567600013688
'                           func_normsdist: time for 120,000 calls:     2.57450999924913E-02
'                           func_normsdist is faster by a factor of:    11.635907419317
'                           func_normsdistdd: time for 120,000 calls:   0.142931300099008
'                           func_normsdistdd2: time for 120,000 calls:  0.250197399989702

'So contrary to 2015 results func_normsdistdd looks to be about twice as fast as calling into
'Application.WorksheetFunction.Norm_S_Dist and func_normsdist looks to be >11 times as fast calling in to Norm_S_Dist
' -----------------------------------------------------------------------------------------------------------------------
Sub TestCumulativeNormalFunctions()
          Dim Res As Double
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double
          Dim t5 As Double
          Dim z As Double
          Const StepSize = 0.0001

1         t1 = sElapsedTime()
2         For z = -6 To 6 Step StepSize
3             Res = Application.WorksheetFunction.Norm_S_Dist(z, True)
4         Next z
5         t2 = sElapsedTime()
6         For z = -6 To 6 Step StepSize
7             Res = func_normsdist(z)
8         Next z
9         t3 = sElapsedTime()
10        For z = -6 To 6 Step StepSize
11            Res = func_normsdistdd(z)
12        Next z
13        t4 = sElapsedTime()
14        For z = -6 To 6 Step StepSize
15            Res = func_normsdistdd2(z)
16        Next z
17        t5 = sElapsedTime()

18        Debug.Print "Application.WorksheetFunction.Norm_S_Dist: time for " + Format$(12 / StepSize, "###,###") + " calls:", t2 - t1
19        Debug.Print "                           func_normsdist: time for " + Format$(12 / StepSize, "###,###") + " calls:", t3 - t2
20        Debug.Print "                           func_normsdist is faster by a factor of: ", (t2 - t1) / (t3 - t2)
21        Debug.Print "                           func_normsdistdd: time for " + Format$(12 / StepSize, "###,###") + " calls:", t4 - t3
22        Debug.Print "                           func_normsdistdd2: time for " + Format$(12 / StepSize, "###,###") + " calls:", t5 - t4
23        Debug.Print String(90, "=")
End Sub

Function func_normsdist(z As Double) As Double
          '******************************************************************
          '*  Adapted from http://lib.stat.cmu.edu/apstat/66
          '******************************************************************

          Const a0 = 0.5
          Const a1 = 0.398942280444
          Const a2 = 0.399903438505
          Const a3 = 5.75885480458
          Const a4 = 29.8213557808
          Const a5 = 2.62433121679
          Const a6 = 48.6959930692
          Const a7 = 5.92885724438

          Const b0 = 0.398942280385
          Const b1 = 3.8052 * 10 ^ (-8)
          Const b2 = 1.00000615302
          Const b3 = 3.98064794 * 10 ^ (-4)
          Const b4 = 1.98615381364
          Const b5 = 0.151679116635
          Const b6 = 5.29330324926
          Const b7 = 4.8385912808
          Const b8 = 15.1508972451
          Const b9 = 0.742380924027
          Const b10 = 30.789933034
          Const b11 = 3.99019417011

          Dim pdf As Double
          Dim q As Double
          Dim Temp As Double
          Dim y As Double
          Dim zabs As Double

1         zabs = Abs(z)

2         If zabs <= 12.7 Then
3             y = a0 * z * z
4             pdf = Exp(-y) * b0
5             If zabs <= 1.28 Then
6                 Temp = y + a3 - a4 / (y + a5 + a6 / (y + a7))
7                 q = a0 - zabs * (a1 - a2 * y / Temp)
8             Else
9                 Temp = (zabs - b5 + b6 / (zabs + b7 - b8 / (zabs + b9 + b10 / (zabs + b11))))
10                q = pdf / (zabs - b1 + (b2 / (zabs + b3 + b4 / Temp)))
11            End If
12        Else
13            pdf = 0
14            q = 0
15        End If

16        If z < 0 Then
17            func_normsdist = q
18        Else
19            func_normsdist = 1 - q
20        End If
End Function

Function func_normsinv(p As Double) As Double
          '***********************************************************
          '*  Adapted from the extremely efficient algorithm found by
          '*  Peter Acklam (http://home.online.no/~pjacklam/notes/invnorm/).
          '*  VBA adaptation by John Herrero http://home.online.no/~pjacklam/notes/invnorm/impl/herrero/inversecdf.txt
          '*  Relative error is less than 1.15* 10^-9 everywhere
          '***********************************************************

          Const a1 = -39.6968302866538
          Const a2 = 220.946098424521
          Const a3 = -275.928510446969
          Const a4 = 138.357751867269
          Const a5 = -30.6647980661472
          Const a6 = 2.50662827745924

          Const b1 = -54.4760987982241
          Const b2 = 161.585836858041
          Const b3 = -155.698979859887
          Const b4 = 66.8013118877197
          Const b5 = -13.2806815528857

          Const C1 = -7.78489400243029E-03
          Const C2 = -0.322396458041136
          Const c3 = -2.40075827716184
          Const c4 = -2.54973253934373
          Const c5 = 4.37466414146497
          Const c6 = 2.93816398269878

          Const d1 = 7.78469570904146E-03
          Const d2 = 0.32246712907004
          Const d3 = 2.445134137143
          Const d4 = 3.75440866190742

          'Define break-points
          Const p_low = 0.02425
          Const p_high = 1 - p_low

          'Define work variables
          Dim q As Double
          Dim R As Double

          'If argument out of bounds, raise error
1         If p <= 0 Or p >= 1 Then Err.Raise 5

2         If p < p_low Then
              'Rational approximation for lower region
3             q = Sqr(-2 * Log(p))
4             func_normsinv = (((((C1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                  ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
5         ElseIf p <= p_high Then
              'Rational approximation for lower region
6             q = p - 0.5
7             R = q * q
8             func_normsinv = (((((a1 * R + a2) * R + a3) * R + a4) * R + a5) * R + a6) * q / _
                  (((((b1 * R + b2) * R + b3) * R + b4) * R + b5) * R + 1)
9         ElseIf p < 1 Then
              'Rational approximation for upper region
10            q = Sqr(-2 * Log(1 - p))
11            func_normsinv = -(((((C1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                  ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
12        End If
End Function

Function func_normsdense(z As Double) As Double
1         func_normsdense = 0.398942280385 * Exp(-(z * z) / 2)
End Function

Function func_binormcum(z1 As Double, z2 As Double, rho As Double)

          Const sqr2pi = 2.506628274631
          Const lowerlimit = -10

          Dim w(1 To 12) As Double
          Dim x(1 To 12)
1         x(1) = 6.40568928626056E-02
2         w(1) = 0.127938195346752
3         x(2) = 0.191118867473616
4         w(2) = 0.125837456346828
5         x(3) = 0.315042679696163
6         w(3) = 0.121670472927803
7         x(4) = 0.433793507626045
8         w(4) = 0.115505668053726
9         x(5) = 0.54542147138884
10        w(5) = 0.107444270115966
11        x(6) = 0.648093651936976
12        w(6) = 9.76186521041137E-02
13        x(7) = 0.740124191578554
14        w(7) = 8.61901615319532E-02
15        x(8) = 0.820001985973903
16        w(8) = 7.33464814110803E-02
17        x(9) = 0.886415527004401
18        w(9) = 5.92985849154367E-02
19        x(10) = 0.938274552002733
20        w(10) = 4.42774388174197E-02
21        x(11) = 0.974728555971309
22        w(11) = 2.85313886289337E-02
23        x(12) = 0.995187219997021
24        w(12) = 1.23412297999865E-02

          Dim i As Long
          Dim integral As Double
          Dim xl As Double
          Dim xm As Double
          Dim xx As Double

25        xm = 0.5 * (z1 + lowerlimit)
26        xl = 0.5 * (z1 - lowerlimit)
27        integral = 0
28        For i = 1 To 12
29            xx = xm + (xl * x(i))
30            integral = integral + w(i) * xl * (1 / sqr2pi) * Exp(-(xx * xx) / 2) * func_normsdist((z2 - rho * xx) / Sqr(1 - rho * rho))
31            xx = xm - (xl * x(i))
32            integral = integral + w(i) * xl * (1 / sqr2pi) * Exp(-(xx * xx) / 2) * func_normsdist((z2 - rho * xx) / Sqr(1 - rho * rho))
33        Next i
34        func_binormcum = integral
End Function

Function func_normsdistdd(z As Double) As Double
          '****************************************************************
          '*     Normal distribution probabilities accurate to 1d-15.
          '*     Reference: J.L. Schonfelder, Math Comp 32(1978), pp 1232-1240.
          '****************************************************************
          Static a(0 To 24) As Double
          Dim b As Double
          Dim bm As Double
          Dim BP As Double
          Dim i As Long
          Dim p As Double
          Dim t As Double
          Dim xa As Double

          Const RTWO = 1.4142135623731

1         If a(0) = 0 Then
2             a(0) = 0.6101430819232
3             a(1) = -0.434841272712578
4             a(2) = 0.176351193643605
5             a(3) = -6.07107956092494E-02
6             a(4) = 1.77120689956941E-02
7             a(5) = -4.32111938556729E-03
8             a(6) = 8.54216676887099E-04
9             a(7) = -1.27155090609163E-04
10            a(8) = 1.12481672436712E-05
11            a(9) = 3.13063885421821E-07
12            a(10) = -2.70988068537762E-07
13            a(11) = 3.07376227014077E-08
14            a(12) = 2.51562038481762E-09
15            a(13) = -1.02892992132032E-09
16            a(14) = 2.99440521199499E-11
17            a(15) = 2.60517896872669E-11
18            a(16) = -2.63483992417197E-12
19            a(17) = -6.43404509890636E-13
20            a(18) = 1.12457401801663E-13
21            a(19) = 1.72815333899861E-14
22            a(20) = -4.26410169494238E-15
23            a(21) = -5.45371977880191E-16
24            a(22) = 1.58697607761671E-16
25            a(23) = 2.0899837844334E-17
26            a(24) = -5.900526869409E-18
27        End If

28        xa = Abs(z) / RTWO
29        If (xa > 100) Then
30            p = 0
31        Else
32            t = (8 * xa - 30) / (4 * xa + 15)
33            bm = 0
34            b = 0
35            For i = 24 To 0 Step -1
36                BP = b
37                b = bm
38                bm = t * b - BP + a(i)
39            Next
40            p = Exp(-xa * xa) * (bm - BP) / 4
41        End If
42        If (z > 0) Then p = 1 - p
43        func_normsdistdd = p
End Function

Function func_normsinvdd(p As Double) As Double
          '*****************************************************************
          '* More accurate func_normsinv, see http://home.online.no/~pjacklam/notes/invnorm/#Overview
          '* depends on the accuracy of func_normsdistdd, which needs to be available
          '*****************************************************************
          Dim ee As Double
          Dim uu As Double
          Dim xx As Double

1         xx = func_normsinv(p)
2         ee = func_normsdistdd(xx) - p
3         uu = ee * 2.506628274631 * Exp(xx * xx / 2)

4         func_normsinvdd = xx - uu / (1 + xx * uu / 2)
End Function

Function func_binormcumdd(dh As Double, DK As Double, R As Double) As Double
          '****************************************************************
          '*     A function for computing bivariate normal probabilities.
          '*
          '*       Alan Genz
          '*       Department of Mathematics
          '*       Washington State University
          '*       Pullman, WA 99164-3113
          '*       Email : alangenz@wsu.edu
          '*
          '*    This function is based on the method described by
          '*    Drezner, Z and G.O. Wesolowsky, (1989),
          '*    On the computation of the bivariate normal integral,
          '*    Journal of Statist. Comput. Simul. 35, pp. 101-107,
          '*    with major modifications for double precision, and for |R| close to 1.
          '*
          '*    func_binormcumdd calculates the probability that X < DH and Y < DK.
          '*
          '*    Parameters
          '*
          '*    DH  DOUBLE PRECISION, integration limit
          '*    DK  DOUBLE PRECISION, integration limit
          '*    R   DOUBLE PRECISION, correlation coefficient
          '*****************************************************************
          Const twopi = 6.28318530717959
          Dim a As Double
          Dim ASR As Double
          Dim ASt As Double
          Dim b As Double
          Dim BS As Double
          Dim BVN As Double
          Dim c As Double
          Dim D As Double
          Dim H As Double
          Dim HK As Double
          Dim Hs As Double
          Dim i As Long
          Dim ISt As Long
          Dim k As Double
          Dim LG As Long
          Dim NG As Long
          Dim RS As Double
          Dim SN As Double
          Dim w(1 To 10, 1 To 3) As Double
          Dim x(1 To 10, 1 To 3) As Double
          Dim Xs As Double
          '*     Gauss Legendre Points and Weights, N =  6
1         w(1, 1) = 0.17132449237917
2         x(1, 1) = -0.932469514203152
3         w(2, 1) = 0.360761573048138
4         x(2, 1) = -0.661209386466265
5         w(3, 1) = 0.46791393457269
6         x(3, 1) = -0.238619186083197
          '*     Gauss Legendre Points and Weights, N = 12
7         w(1, 2) = 4.71753363865118E-02
8         x(1, 2) = -0.981560634246719
9         w(2, 2) = 0.106939325995318
10        x(2, 2) = -0.904117256370475
11        w(3, 2) = 0.160078328543346
12        x(3, 2) = -0.769902674194305
13        w(4, 2) = 0.203167426723066
14        x(4, 2) = -0.587317954286617
15        w(5, 2) = 0.233492536538355
16        x(5, 2) = -0.36783149899818
17        w(6, 2) = 0.249147045813403
18        x(6, 2) = -0.125233408511469
          '*     Gauss Legendre Points and Weights, N = 20
19        w(1, 3) = 1.76140071391521E-02
20        x(1, 3) = -0.993128599185095
21        w(2, 3) = 4.06014298003869E-02
22        x(2, 3) = -0.963971927277914
23        w(3, 3) = 6.26720483341091E-02
24        x(3, 3) = -0.912234428251326
25        w(4, 3) = 8.32767415767048E-02
26        x(4, 3) = -0.839116971822219
27        w(5, 3) = 0.10193011981724
28        x(5, 3) = -0.746331906460151
29        w(6, 3) = 0.118194531961518
30        x(6, 3) = -0.636053680726515
31        w(7, 3) = 0.131688638449177
32        x(7, 3) = -0.510867001950827
33        w(8, 3) = 0.142096109318382
34        x(8, 3) = -0.37370608871542
35        w(9, 3) = 0.149172986472604
36        x(9, 3) = -0.227785851141645
37        w(10, 3) = 0.152753387130726
38        x(10, 3) = -7.65265211334973E-02

39        If Abs(R) < 0.3 Then
40            NG = 1
41            LG = 3
42        ElseIf Abs(R) < 0.75 Then
43            NG = 2
44            LG = 6
45        Else
46            NG = 3
47            LG = 10
48        End If

49        H = -dh
50        k = -DK
51        HK = H * k
52        BVN = 0
53        If Abs(R) < 0.925 Then
54            If Abs(R) > 0 Then
55                Hs = (H * H + k * k) / 2
56                ASR = Atn(R / Sqr(1 - R * R))
57                For i = 1 To LG
58                    For ISt = -1 To 1 Step 2
59                        SN = Sin(ASR * (ISt * x(i, NG) + 1) / 2)
60                        BVN = BVN + w(i, NG) * Exp((SN * HK - Hs) / (1 - SN * SN))
61                    Next
62                Next
63                BVN = BVN * ASR / (2 * twopi)
64            End If
65            BVN = BVN + func_normsdistdd(-H) * func_normsdistdd(-k)
66        Else
67            If R < 0 Then
68                k = -k
69                HK = -HK
70            End If
71            If Abs(R) < 1 Then
72                ASt = (1 - R) * (1 + R)
73                a = Sqr(ASt)
74                BS = (H - k) * (H - k)
75                c = (4 - HK) / 8
76                D = (12 - HK) / 16
77                ASR = -(BS / ASt + HK) / 2
78                If ASR > -100 Then BVN = a * Exp(ASR) * (1 - c * (BS - ASt) * (1 - D * BS / 5) / 3 + c * D * ASt * ASt / 5)
79                If -HK < 100 Then
80                    b = Sqr(BS)
81                    BVN = BVN - Exp(-HK / 2) * Sqr(twopi) * func_normsdistdd(-b / a) * b * (1 - c * BS * (1 - D * BS / 5) / 3)
82                End If
83                a = a / 2
84                For i = 1 To LG
85                    For ISt = -1 To 1 Step 2
86                        Xs = (a * (ISt * x(i, NG) + 1)) ^ 2
87                        RS = Sqr(1 - Xs)
88                        ASR = -(BS / Xs + HK) / 2
89                        If ASR > -100 Then
90                            BVN = BVN + a * w(i, NG) * Exp(ASR) * (Exp(-HK * (1 - RS) / (2 * (1 + RS))) / RS - (1 + c * Xs * (1 + D * Xs)))
91                        End If
92                    Next
93                Next
94                BVN = -BVN / twopi
95            End If
96            If R > 0 Then
97                If k > H Then
98                    BVN = BVN + func_normsdistdd(-k)
99                Else
100                   BVN = BVN + func_normsdistdd(-H)
101               End If
102           Else
103               BVN = -BVN
104               If k > H Then BVN = BVN + func_normsdistdd(k) - func_normsdistdd(H)
105           End If
106       End If
107       func_binormcumdd = BVN
End Function

Function func_normsdistdd2(z As Double)
          '******************************************************************************************
          '*
          '*  Highly accurate normsdist function.
          '*  Adapted from FORTRAN routine published on: http://people.sc.fsu.edu/~burkardt/f77_src/specfun/specfun.f
          '*  by William Cody
          '*
          '*  Reference:
          '*
          '*    William Cody,
          '*    Rational Chebyshev Approximations for the Error Function,
          '*    Mathematics of Computation,
          '*    Volume 23, Number 107, July 1969, pages 631-638.
          '*
          '*  Appears slightly slower than the routine func_normsdistdd, so not the first choice
          '******************************************************************************************

          Dim a(1 To 5) As Double
          Dim b(1 To 4) As Double
          Dim c(1 To 9) As Double
          Dim D(1 To 8) As Double
          Dim i As Long
          Dim p(1 To 6) As Double
          Dim q(1 To 5) As Double
          Dim Result As Double
          Dim x As Double
          Dim xden As Double
          Dim xnum As Double
          Dim y As Double
          Dim ysq As Double

          Const sqrpi = 0.564189583547756
          Const thresh = 0.46875

          Const xsmall = 1.11E-16
          Const xbig = 26.543

1         a(1) = 3.16112374387057
2         a(2) = 113.86415415105
3         a(3) = 377.485237685302
4         a(4) = 3209.37758913847
5         a(5) = 0.185777706184603

6         b(1) = 23.6012909523441
7         b(2) = 244.024637934444
8         b(3) = 1282.61652607737
9         b(4) = 2844.23683343917

10        c(1) = 0.56418849698867
11        c(2) = 8.88314979438838
12        c(3) = 66.1191906371416
13        c(4) = 298.6351381974
14        c(5) = 881.952221241769
15        c(6) = 1712.04761263407
16        c(7) = 2051.07837782607
17        c(8) = 1230.339354798
18        c(9) = 2.15311535474404E-08

19        D(1) = 15.7449261107098
20        D(2) = 117.693950891312
21        D(3) = 537.18110186201
22        D(4) = 1621.38957456669
23        D(5) = 3290.79923573346
24        D(6) = 4362.61909014325
25        D(7) = 3439.36767414372
26        D(8) = 1230.33935480375

27        p(1) = 0.305326634961232
28        p(2) = 0.360344899949804
29        p(3) = 0.125781726111229
30        p(4) = 1.60837851487423E-02
31        p(5) = 6.58749161529838E-04
32        p(6) = 1.63153871373021E-02

33        q(1) = 2.56852019228982
34        q(2) = 1.87295284992346
35        q(3) = 0.527905102951428
36        q(4) = 6.05183413124413E-02
37        q(5) = 2.33520497626869E-03

38        x = -z / 1.4142135623731
39        y = Abs(x)

40        If y <= thresh Then
41            ysq = 0
42            If xsmall < y Then
43                ysq = y ^ 2
44            End If
45            xnum = a(5) * ysq
46            xden = ysq
47            For i = 1 To 3
48                xnum = (xnum + a(i)) * ysq
49                xden = (xden + b(i)) * ysq
50            Next
51            Result = 0.5 - 0.5 * x * (xnum + a(4)) / (xden + b(4))
52        ElseIf y <= 4 Then
53            xnum = c(9) * y
54            xden = y
55            For i = 1 To 7
56                xnum = (xnum + c(i)) * y
57                xden = (xden + D(i)) * y
58            Next
59            Result = (xnum + c(8)) / (xden + D(8))
60            Result = 0.5 * Result * Exp(-y ^ 2)
61            If x < 0 Then
62                Result = 1 - Result
63            End If
64        Else
65            Result = 0
66            If y >= xbig Then GoTo EndOfFunction
67            ysq = 1 / (y * y)
68            xnum = p(6) * ysq
69            xden = ysq
70            For i = 1 To 4
71                xnum = (xnum + p(i)) * ysq
72                xden = (xden + q(i)) * ysq
73            Next

74            Result = ysq * (xnum + p(5)) / (xden + q(5))
75            Result = (sqrpi - Result) / y
76            Result = 0.5 * Result * Exp(-y ^ 2)
77            If x < 0 Then
78                Result = 1 - Result
79            End If
80        End If
EndOfFunction:
81        func_normsdistdd2 = Result
End Function

Function func_binormdense(z1 As Double, z2 As Double, rho As Double)
          'bivariate normal density function

1         func_binormdense = 0.159154943091895 / Sqr(1 - rho ^ 2) * Exp(-0.5 / (1 - rho ^ 2) * (z1 ^ 2 + z2 ^ 2 - 2 * rho * z1 * z2))
End Function

Function func_logdist(Arg As Double, mu As Double, Ss As Double) As Double
          '****************************************************************************
          '* Lognormal distribution: support [0,+inf[, mean= exp(mu+ss^2/2),
          '* Stdev^2 = (exp(ss^2)-1)*exp(2*mu+ss^2)
          '****************************************************************************
1         func_logdist = func_normsdist((Log(Arg) - mu) / Ss)
End Function

Function func_loginv(Arg As Double, mu As Double, Ss As Double) As Double
1         func_loginv = Exp(mu + Ss * func_normsinv(Arg))
End Function

Function func_logdense(z As Double, mean As Double, StDev As Double) As Double
1         If z <= 0 Then
2             func_logdense = 0
3         Else
4             func_logdense = 0.398942280401433 / z / StDev * Exp(-(Log(z) - mean) ^ 2 / 2 / StDev / StDev)
5         End If
End Function

Function func_logdistdd(Arg As Double, mu As Double, Ss As Double) As Double
1         func_logdistdd = func_normsdistdd((Log(Arg) - mu) / Ss)
End Function

Function func_loginvdd(Arg As Double, mu As Double, Ss As Double) As Double
1         func_loginvdd = Exp(mu + Ss * func_normsinvdd(Arg))
End Function
