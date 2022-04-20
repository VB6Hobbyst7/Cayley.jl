Attribute VB_Name = "modMinPackSolver"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modMinPackSolver
' Author    : Philip Swannell
' Date      : 30-Apr-2015
' Purpose   : Found this multi-dimensional solver code at http://www.quantcode.com/modules/mydownloads/singlefile.php?lid=511
'             The code purports to be a port to VB of MINPACK - see http://en.wikipedia.org/wiki/MINPACK
'             Modifications made:
'        1)   Declare two variables tmp in method fsolve and TEMP2 in method fcntest, so that code compiles with Option Explicit.
'        2)   Removed declaration of a number of variables that MZTools reveals not to be used.
'        3)   Edited method fcn to use direct calls to the objective function rather than calling an arbitrary function
'             via Application.Run. Application.Run breaks the error handling stack so that run-time error dialogs pop up
'             as code is running :-(
'        4)   Switched variables fvec, WA1 and WA4 to be arrays of doubles not arrays of variants. This fixes a memory leak! See code comment
'             this changes are commented with "Explicit double array PB#20150618"
'             Option Private Module to make functions not visible from Excel
'        5)   Overwrote change 4) above by using Dim in the line before Redim is first called. Change suggested by Rubberduck 11-May-2019
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private FunctionName_ As String
Option Private Module

'Minpack Copyright Notice (1999) University of Chicago.  All rights reserved

'Redistribution and use in source and binary forms, with or
'without modification, are permitted provided that the
'following conditions are met:

'1. Redistributions of source code must retain the above
'copyright notice, this list of conditions and the following
'disclaimer.

'2. Redistributions in binary form must reproduce the above
'copyright notice, this list of conditions and the following
'disclaimer in the documentation and/or other materials
'provided with the distribution.

'3. The end-user documentation included with the
'redistribution, if any, must include the following
'acknowledgment:

'   "This product includes software developed by the
'   University of Chicago, as Operator of Argonne National
'   Laboratory.

'Alternately, this acknowledgment may appear in the software
'itself, if and wherever such third-party acknowledgments
'normally appear.

'4. WARRANTY DISCLAIMER. THE SOFTWARE IS SUPPLIED "AS IS"
'WITHOUT WARRANTY OF ANY KIND. THE COPYRIGHT HOLDER, THE
'UNITED STATES, THE UNITED STATES DEPARTMENT OF ENERGY, AND
'THEIR EMPLOYEES: (1) DISCLAIM ANY WARRANTIES, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO ANY IMPLIED WARRANTIES
'OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE
'OR NON-INFRINGEMENT, (2) DO NOT ASSUME ANY LEGAL LIABILITY
'OR RESPONSIBILITY FOR THE ACCURACY, COMPLETENESS, OR
'USEFULNESS OF THE SOFTWARE, (3) DO NOT REPRESENT THAT USE OF
'THE SOFTWARE WOULD NOT INFRINGE PRIVATELY OWNED RIGHTS, (4)
'DO NOT WARRANT THAT THE SOFTWARE WILL FUNCTION
'UNINTERRUPTED, THAT IT IS ERROR-FREE OR THAT ANY ERRORS WILL
'BE CORRECTED.

'5. LIMITATION OF LIABILITY. IN NO EVENT WILL THE COPYRIGHT
'HOLDER, THE UNITED STATES, THE UNITED STATES DEPARTMENT OF
'ENERGY, OR THEIR EMPLOYEES: BE LIABLE FOR ANY INDIRECT,
'INCIDENTAL, CONSEQUENTIAL, SPECIAL OR PUNITIVE DAMAGES OF
'ANY KIND OR NATURE, INCLUDING BUT NOT LIMITED TO LOSS OF
'PROFITS OR LOSS OF DATA, FOR ANY REASON WHATSOEVER, WHETHER
'SUCH LIABILITY IS ASSERTED ON THE BASIS OF CONTRACT, TORT
'(INCLUDING NEGLIGENCE OR STRICT LIABILITY), OR OTHERWISE,
'EVEN IF ANY OF SAID PARTIES HAS BEEN WARNED OF THE
'POSSIBILITY OF SUCH LOSS OR DAMAGES.

Public Function fsolve(FunctionName As String, xguessvec As Variant) As Variant
1         FunctionName_ = FunctionName
          Dim j As Single
          Dim N As Single
          Dim tmp As Variant
2         On Error GoTo err_sub
3         N = UBound(xguessvec)
4         GoTo no_error
err_sub:
5         N = 1
6         tmp = xguessvec
7         ReDim xguessvec(1 To 1)
8         xguessvec(1) = tmp
no_error:
          Dim EPSFCN As Double
          Dim Factor As Double
          Dim Info As Single
          Dim LDFJAC As Single
          Dim LR As Single
          Dim MAXFEV As Single
          Dim ML As Single
          Dim Mode As Single
          Dim mu As Single
          Dim NFEV As Single
          Dim NPRINT As Single
          Dim NWRITE As Single
          Dim XTOL As Double
          '          Dim FNORM As Double
          Dim x() As Double                     ' Added by PGS 11-May-2019
9         ReDim x(1 To N)
          Dim fvec() As Double                  ' Added by PGS 11-May-2019
10        ReDim fvec(1 To N)
          Dim diag() As Double                  ' Added by PGS 11-May-2019
11        ReDim diag(1 To N)
          Dim fjac() As Double                  ' Added by PGS 11-May-2019
12        ReDim fjac(1 To N, 1 To N)
13        LR = (N * (N + 1)) / 2
          Dim R() As Double                     ' Added by PGS 11-May-2019
14        ReDim R(1 To LR)
          Dim QTF() As Double                   ' Added by PGS 11-May-2019
15        ReDim QTF(1 To N)
          Dim WA1() As Double                   ' Added by PGS 11-May-2019
16        ReDim WA1(1 To N)
          Dim WA2() As Double                   ' Added by PGS 11-May-2019
17        ReDim WA2(1 To N)
          Dim WA3() As Double                   ' Added by PGS 11-May-2019
18        ReDim WA3(1 To N)
          Dim WA4() As Double                   ' Added by PGS 11-May-2019
19        ReDim WA4(1 To N)

20        NWRITE = 6
          'N = 9
          '
          '     THE FOLLOWING STARTING VALUES PROVIDE A ROUGH SOLUTION.
          '
21        For j = 1 To N
22            x(j) = xguessvec(j)
23        Next j
          '
24        LDFJAC = N

          '
          '     SET XTOL TO THE SQUARE ROOT OF THE MACHINE PRECISION.
          '     UNLESS HIGH PRECISION SOLUTIONS ARE REQUIRED,
          '     THIS IS THE RECOMMENDED SETTING.
          '
          'XTOL = 0.0001
25        XTOL = 0.000000000000001
26        MAXFEV = 2000
27        ML = 1
          'MU = 1
28        mu = N - 1
29        EPSFCN = 0
30        Mode = 2
31        For j = 1 To N
32            diag(j) = 1
33        Next j
34        Factor = 100
35        NPRINT = 0
          Dim fcn2 As String
36        fcn2 = "dummy"
37        hybrd fcn2, N, x, fvec, XTOL, MAXFEV, ML, mu, EPSFCN, diag, _
              Mode, Factor, NPRINT, Info, NFEV, fjac, LDFJAC, _
              R, LR, QTF, WA1, WA2, WA3, WA4

38        fsolve = x
          '      iflag = 1
          '     fcn N, x, fvec, iflag
          '    FNORM = myenorm(N, fvec)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestOverheadOfApplicationDotRun
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Test my theory that as well as breaking the error handling stack, Application.Run is also slow...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestOverheadOfApplicationDotRun()
          Dim i As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Const NumCalls = 10000
1         t1 = sElapsedTime()
2         For i = 1 To NumCalls
3             Res = SafeMin(1, 2)
4         Next i
5         t2 = sElapsedTime()
6         For i = 1 To NumCalls
7             Res = Application.Run("SafeMin", 1, 2)
8         Next i
9         t3 = sElapsedTime()

10        Debug.Print "Direct", t2 - t1, "Via Application.Run", t3 - t2, "Ratio", (t3 - t2) / (t2 - t1)
End Sub

Public Sub fcn(ByRef N As Single, ByRef x As Variant, ByRef fvec() As Double, _
        ByRef iflag As Variant)

          ' Explicit double array fvec PB#20150618

1         If (iflag <> 0) Then GoTo L5
2         Exit Sub
L5:
          'Edited by Philip Swannell 14-May-2015. Use of Application.Run is _
           perhaps slow, but certainly breaks the error handling stack, so we _
           should edit this function each time we use the fsolve method...

3         If FunctionName_ = "BasketSolveObjectiveFn" Then
4             If N = 1 Then
                  'fvec(1) = BasketSolveObjectiveFn(x(1))
5             Else
6                 fvec = BasketSolveObjectiveFn(x)
7             End If
8         ElseIf FunctionName_ = "OptSolveVolObjectiveFn" Then
9             If N = 1 Then
10                fvec(1) = OptSolveVolObjectiveFn(x(1))
11            Else
12                fvec = OptSolveVolObjectiveFn(x)
13            End If
14        Else
15            If N = 1 Then
16                fvec(1) = Application.Run(FunctionName_, x(1))
17            Else
18                fvec = Application.Run(FunctionName_, x)
19            End If
20        End If
End Sub
Public Sub fcntest(ByRef N As Single, ByRef x As Variant, ByRef fvec As Variant, _
        ByRef iflag As Variant)
          '         calculate the functions at x and
          '         return this vector in fvec.
          Dim k As Single
          Dim Temp As Double
          Dim temp1 As Double
          Dim TEMP2 As Double

          '       DOUBLE PRECISION ONE,TEMP,TEMP1,TEMP2,THREE,TWO,ZERO
          '       DATA ZERO,ONE,TWO,THREE /0.D0,1.D0,2.D0,3.D0/
          '
1         If (iflag <> 0) Then GoTo L5
          '
          '     INSERT PRINT STATEMENTS HERE WHEN NPRINT IS POSITIVE.
          '
2         Return
L5:
3         For k = 1 To N
4             Temp = (3 - 2 * x(k)) * x(k)
5             temp1 = 0
6             If (k <> 1) Then temp1 = x(k - 1)
7             TEMP2 = 0
8             If (k <> N) Then TEMP2 = x(k + 1)
9             fvec(k) = Temp - temp1 - 2 * TEMP2 + 1
10        Next k
End Sub

Public Sub hybrd(fcn2 As String, ByRef N As Single, ByRef x As Variant, ByRef fvec() As Double, _
        ByRef XTOL As Double, ByRef MAXFEV As Single, ByRef ML As Single, ByRef mu As Single, _
        ByRef EPSFCN As Double, ByRef diag As Variant, _
        ByRef Mode As Single, ByRef Factor As Double, ByRef NPRINT As Single, ByRef Info As Single, _
        ByRef NFEV As Single, ByRef fjac As Variant, ByRef LDFJAC As Single, ByRef R As Variant, _
        ByRef LR As Single, ByRef QTF As Variant, ByRef WA1() As Double, _
        ByRef WA2 As Variant, ByRef WA3 As Variant, ByRef WA4() As Double)

          ' Explicit double array WA1 PB#20150618
          ' Explicit double array WA4 PB#20150618
          ' Explicit double array fvec PB#20150618

          'integer n,maxfev,ml,mu,mode,nprint,info,nfev,ldfjac,lr
          'double precision xtol,epsfcn,factor
          'double precision x(n),fvec(n),diag(n),fjac(ldfjac,n),r(lr),
          '                 qtf(n),wa1(n),wa2(n),wa3(n),wa4(n)

          ' external fcn
          Dim i As Single
          Dim iflag As Single
          Dim iter As Single
          Dim IWA() As Double
          Dim j As Single
          Dim jm1 As Single
          Dim L As Single
          Dim msum As Single
          Dim ncfail As Single
          Dim ncsuc As Single
          Dim nslow1 As Single
          Dim nslow2 As Single
1         ReDim IWA(1 To 1)
          Dim jeval As Boolean
          Dim sing As Boolean

          'integer i,iflag,iter,j,jm1,l,msum,ncfail,ncsuc,nslow1,nslow2
          'integer iwa(1)
          'logical jeval, sing
          Dim actred As Double
          Dim delta As Double
          Dim epsmch As Double
          Dim FNORM As Double
          Dim fnorm1 As Double
          Dim one As Double
          Dim p0001 As Double
          Dim p001 As Double
          Dim P1 As Double
          Dim p5 As Double
          Dim pnorm As Double
          Dim prered As Double
          Dim Ratio As Double
          Dim sum As Double
          Dim Temp As Double
          Dim xnorm As Double
          Dim zero As Double

          'double precision actred,delta,epsmch,fnorm,fnorm1,one,pnorm,
          '*                 prered,p1,p5,p001,p0001,ratio,sum,temp,xnorm,
          '*                 zero
          'double precision dpmpar,enorm
2         one = 1
3         P1 = 0.1
4         p5 = 0.5
5         p001 = 0.001
6         p0001 = 0.0001
7         zero = 0

          'data one, p1, p5, p001, p0001, zero
          '*     /1.0d0,1.0d-1,5.0d-1,1.0d-3,1.0d-4,0.0d0/
          '
          '     epsmch is the machine precision.
          '
8         epsmch = 1.2E-16
          '
9         Info = 0
10        iflag = 0
11        NFEV = 0
          '
          '     check the input parameters for errors.
          '
12        If ((N <= 0) Or (XTOL < zero) Or (MAXFEV <= 0) _
              Or (ML < 0) Or (mu < 0) Or (Factor <= zero) _
              Or (LDFJAC < N) Or (LR < (N * (N + 1)) / 2)) Then GoTo L300

13        If (Mode <> 2) Then GoTo L20
14        For j = 1 To N
15            If (diag(j) <= zero) Then GoTo L300
16        Next j
L20:
          '
          '     evaluate the function at the starting point
          '     and calculate its norm.
          '
17        iflag = 1
18        fcn N, x, fvec, iflag
19        NFEV = 1
20        If (iflag < 0) Then GoTo L300
21        FNORM = myenorm(N, fvec)
          '
          '     determine the number of calls to fcn needed to compute
          '     the jacobian matrix.
          '
22        msum = min0(ML + mu + 1, N)
          '
          '     initialize iteration counter and monitors.
          '
23        iter = 1
24        ncsuc = 0
25        ncfail = 0
26        nslow1 = 0
27        nslow2 = 0
          '
          '     beginning of the outer loop.
          '
L30:
28        jeval = True
          '
          '        calculate the jacobian matrix.
          '
29        iflag = 2
30        fdjac1 fcn2, N, x, fvec, fjac, LDFJAC, iflag, ML, mu, EPSFCN, WA1, WA2
31        NFEV = NFEV + msum
32        If (iflag < 0) Then GoTo L300
          '
          '        compute the qr factorization of the jacobian.
          '
33        qrfac N, N, fjac, LDFJAC, False, IWA, 1, WA1, WA2, WA3
          '
          '        on the first iteration and if mode is 1, scale according
          '        to the norms of the columns of the initial jacobian.
          '
34        If (iter <> 1) Then GoTo L70
35        If (Mode = 2) Then GoTo L50
36        For j = 1 To N
37            diag(j) = WA2(j)
38            If (WA2(j) = zero) Then diag(j) = one
39        Next j
L50:
          '
          '        on the first iteration, calculate the norm of the scaled x
          '        and initialize the step bound delta.
          '
40        For j = 1 To N
41            WA3(j) = diag(j) * x(j)
42        Next j
43        xnorm = myenorm(N, WA3)
44        delta = Factor * xnorm
45        If (delta = zero) Then delta = Factor
L70:
          '
          '        form (q transpose)*fvec and store in qtf.
          '
46        For i = 1 To N
47            QTF(i) = fvec(i)
48        Next i
49        For j = 1 To N
50            If (fjac(j, j) = zero) Then GoTo L110
51            sum = zero
52            For i = j To N
53                sum = sum + fjac(i, j) * QTF(i)
54            Next i
55            Temp = -sum / fjac(j, j)
56            For i = j To N
57                QTF(i) = QTF(i) + fjac(i, j) * Temp
58            Next i
L110:
59        Next j
          '
          '        copy the triangular factor of the qr factorization into r.
          '
60        sing = False
61        For j = 1 To N
62            L = j
63            jm1 = j - 1
64            If (jm1 < 1) Then GoTo L140
65            For i = 1 To jm1
66                R(L) = fjac(i, j)
67                L = L + N - i
68            Next i
L140:
69            R(L) = WA1(j)
70            If (WA1(j) = zero) Then sing = True
71        Next j
          '
          '        accumulate the orthogonal factor in fjac.
          '
72        qform N, N, fjac, LDFJAC, WA1
          '
          '        rescale if necessary.
          '
73        If (Mode = 2) Then GoTo L170
74        For j = 1 To N
75            diag(j) = dmax1(diag(j), WA2(j))
76        Next j
L170:
          '
          '        beginning of the inner loop.
          '
L180:
          '
          '           if requested, call fcn to enable printing of iterates.
          '
77        If (NPRINT <= 0) Then GoTo L190
78        iflag = 0
79        If (mymod(iter - 1, NPRINT) = 0) Then
80            fcn N, x, fvec, iflag
81        End If
82        If (iflag < 0) Then GoTo L300
L190:
          '
          '           determine the direction p.
          '
83        dogleg N, R, LR, diag, QTF, delta, WA1, WA2, WA3
          '
          '           store the direction p and x + p. calculate the norm of p.
          '
84        For j = 1 To N
85            WA1(j) = -WA1(j)
86            WA2(j) = x(j) + WA1(j)
87            WA3(j) = diag(j) * WA1(j)
88        Next j
89        pnorm = myenorm(N, WA3)
          '
          '           on the first iteration, adjust the initial step bound.
          '
90        If (iter = 1) Then delta = dmin1(delta, pnorm)
          '
          '           evaluate the function at x + p and calculate its norm.
          '
91        iflag = 1
92        fcn N, WA2, WA4, iflag
93        NFEV = NFEV + 1
94        If (iflag < 0) Then GoTo L300
95        fnorm1 = myenorm(N, WA4)
          '
          '           compute the scaled actual reduction.
          '
96        actred = -one
97        If (fnorm1 < FNORM) Then actred = one - (fnorm1 / FNORM) ^ 2
          '
          '           compute the scaled predicted reduction.
          '
98        L = 1
99        For i = 1 To N
100           sum = zero
101           For j = i To N
102               sum = sum + R(L) * WA1(j)
103               L = L + 1
104           Next j
105           WA3(i) = QTF(i) + sum
106       Next i
107       Temp = myenorm(N, WA3)
108       prered = zero
109       If (Temp < FNORM) Then prered = one - (Temp / FNORM) ^ 2
          '
          '           compute the ratio of the actual to the predicted
          '           reduction.
          '
110       Ratio = zero
111       If (prered > zero) Then Ratio = actred / prered
          '
          '           update the step bound.
          '
112       If (Ratio >= P1) Then GoTo L230
113       ncsuc = 0
114       ncfail = ncfail + 1
115       delta = p5 * delta
116       GoTo L240
L230:
117       ncfail = 0
118       ncsuc = ncsuc + 1
119       If ((Ratio >= p5) Or (ncsuc > 1)) Then delta = dmax1(delta, pnorm / p5)
120       If (dabs(Ratio - one) <= P1) Then delta = pnorm / p5
L240:
          '
          '           test for successful iteration.
          '
121       If (Ratio < p0001) Then GoTo L260
          '
          '           successful iteration. update x, fvec, and their norms.
          '
122       For j = 1 To N
123           x(j) = WA2(j)
124           WA2(j) = diag(j) * x(j)
125           fvec(j) = WA4(j)
126       Next j
127       xnorm = myenorm(N, WA2)
128       FNORM = fnorm1
129       iter = iter + 1
L260:
          '
          '           determine the progress of the iteration.
          '
130       nslow1 = nslow1 + 1
131       If (actred >= p001) Then nslow1 = 0
132       If (jeval) Then nslow2 = nslow2 + 1
133       If (actred >= P1) Then nslow2 = 0
          '
          '           test for convergence.
          '
134       If ((delta <= XTOL * xnorm) Or (FNORM = zero)) Then Info = 1
135       If (Info <> 0) Then GoTo L300
          '
          '           tests for termination and stringent tolerances.
          '
136       If (NFEV >= MAXFEV) Then Info = 2
137       If (P1 * dmax1(P1 * delta, pnorm) <= epsmch * xnorm) Then Info = 3
138       If (nslow2 = 5) Then Info = 4
139       If (nslow1 = 10) Then Info = 5
140       If (Info <> 0) Then GoTo L300
          '
          '           criterion for recalculating jacobian approximation
          '           by forward differences.
          '
141       If (ncfail = 2) Then GoTo L290
          '
          '           calculate the rank one modification to the jacobian
          '           and update qtf if necessary.
          '
142       For j = 1 To N
143           sum = zero
144           For i = 1 To N
145               sum = sum + fjac(i, j) * WA4(i)
146           Next i
147           WA2(j) = (sum - WA3(j)) / pnorm
148           WA1(j) = diag(j) * ((diag(j) * WA1(j)) / pnorm)
149           If (Ratio >= p0001) Then QTF(j) = sum
150       Next j
          '
          '           compute the qr factorization of the updated jacobian.
          '
151       r1updt N, N, R, LR, WA1, WA2, WA3, sing
152       r1mpyq N, N, fjac, LDFJAC, WA2, WA3
153       r1mpyqforqtf 1, N, QTF, 1, WA2, WA3
          '
          '           end of the inner loop.
          '
154       jeval = False
155       GoTo L180
L290:
          '
          '        end of the outer loop.
          '
156       GoTo L30
L300:
          '
          '     termination, either normal or user imposed.
          '
157       If (iflag < 0) Then Info = iflag
158       iflag = 0
159       If (NPRINT > 0) Then fcn N, x, fvec, iflag
End Sub

Private Sub r1updt(ByRef M As Single, ByRef N As Single, ByRef S As Variant, _
        ByRef ls As Single, ByRef u As Variant, ByRef v As Variant, _
        ByRef w As Variant, ByRef sing As Boolean)

          'integer m,n,ls
          'logical sing
          'double precision s(ls),u(m),v(n),w(m)
          Dim i As Single
          Dim j As Single
          Dim jj As Single
          Dim L As Single
          Dim nm1 As Single
          Dim nmj As Single
          'integer i,j,jj,l,nmj,nm1
          Dim giant As Double
          Dim mycos As Double
          Dim mycotan As Double
          Dim mysin As Double
          Dim mytan As Double
          Dim one As Double
          Dim p25 As Double
          Dim p5 As Double
          Dim Tau As Double
          Dim Temp As Double
          Dim zero As Double
          'double precision cos,cotan,giant,one,p5,p25,sin,tan,tau,temp,
          '*                 zero
          'double precision dpmpar
1         one = 1
2         p5 = 0.5
3         p25 = 0.25
4         zero = 0
          'data one,p5,p25,zero /1.0d0,5.0d-1,2.5d-1,0.0d0/
          '
          '     giant is the largest magnitude.
          '
5         giant = 2 ^ 55
          '
          '     initialize the diagonal element pointer.
          '
6         jj = (N * (2 * M - N + 1)) / 2 - (M - N)
          '
          '     move the nontrivial part of the last column of s into w.
          '
7         L = jj
8         For i = N To M
9             w(i) = S(L)
10            L = L + 1
11        Next i
          '
          '     rotate the vector v into a multiple of the n-th unit vector
          '     in such a way that a spike is introduced into w.
          '
12        nm1 = N - 1
13        If (nm1 < 1) Then GoTo L70
14        For nmj = 1 To nm1
15            j = N - nmj
16            jj = jj - (M - j + 1)
17            w(j) = zero
18            If (v(j) = zero) Then GoTo L50
              '
              '        determine a givens rotation which eliminates the
              '        j-th element of v.
              '
19            If (dabs(v(N)) >= dabs(v(j))) Then GoTo L20
20            mycotan = v(N) / v(j)
21            mysin = p5 / dsqrt(p25 + p25 * mycotan ^ 2)
22            mycos = mysin * mycotan
23            Tau = one
24            If (dabs(mycos) * giant > one) Then Tau = one / mycos
25            GoTo L30
L20:
26            mytan = v(j) / v(N)
27            mycos = p5 / dsqrt(p25 + p25 * mytan ^ 2)
28            mysin = mycos * mytan
29            Tau = mysin
L30:
              '
              '        apply the transformation to v and store the information
              '        necessary to recover the givens rotation.
              '
30            v(N) = mysin * v(j) + mycos * v(N)
31            v(j) = Tau
              '
              '        apply the transformation to s and extend the spike in w.
              '
32            L = jj
33            For i = j To M
34                Temp = mycos * S(L) - mysin * w(i)
35                w(i) = mysin * S(L) + mycos * w(i)
36                S(L) = Temp
37                L = L + 1
38            Next i
L50:
39        Next nmj
L70:
          '
          '     add the spike from the rank 1 update to w.
          '
40        For i = 1 To M
41            w(i) = w(i) + v(N) * u(i)
42        Next i
          '
          '     eliminate the spike.
          '
43        sing = False
44        If (nm1 < 1) Then GoTo L140
45        For j = 1 To nm1
46            If (w(j) = zero) Then GoTo L120
              '
              '        determine a givens rotation which eliminates the
              '        j-th element of the spike.
              '
47            If (dabs(S(jj)) >= dabs(w(j))) Then GoTo L90
48            mycotan = S(jj) / w(j)
49            mysin = p5 / dsqrt(p25 + p25 * mycotan ^ 2)
50            mycos = mysin * mycotan
51            Tau = one
52            If (dabs(mycos) * giant > one) Then Tau = one / mycos
53            GoTo L100
L90:
54            mytan = w(j) / S(jj)
55            mycos = p5 / dsqrt(p25 + p25 * mytan ^ 2)
56            mysin = mycos * mytan
57            Tau = mysin
L100:
              '
              '        apply the transformation to s and reduce the spike in w.
              '
58            L = jj
59            For i = j To M
60                Temp = mycos * S(L) + mysin * w(i)
61                w(i) = -mysin * S(L) + mycos * w(i)
62                S(L) = Temp
63                L = L + 1
64            Next i
              '
              '        store the information necessary to recover the
              '        givens rotation.
              '
65            w(j) = Tau
L120:
              '
              '        test for zero diagonal elements in the output s.
              '
66            If (S(jj) = zero) Then sing = True
67            jj = jj + (M - j + 1)
68        Next j
L140:
          '
          '     move w back into the last column of the output s.
          '
69        L = jj
70        For i = N To M
71            S(L) = w(i)
72            L = L + 1
73        Next i
74        If (S(jj) = zero) Then sing = True
End Sub

Private Function min1single(a As Single, b As Single)
1         If a <= b Then
2             min1single = a
3         Else
4             min1single = b
5         End If
End Function
Private Sub qrfac(ByRef M As Single, ByRef N As Single, ByRef a As Variant, ByRef lda As Single, _
        ByRef pivot As Boolean, ByRef ipvt As Variant, ByRef lipvt As Single, ByRef rdiag As Variant, _
        ByRef acnorm As Variant, ByRef wa As Variant)
          Dim ajnorm As Double
          Dim enorm As Double
          Dim epsmch As Double
          Dim i As Single
          Dim j As Single
          Dim jp1 As Single
          Dim k As Single
          Dim kmax As Single
          Dim minmn As Single
          Dim one As Double
          Dim p05 As Double
          Dim sum As Double
          Dim Temp As Double
          Dim zero As Double
1         one = 1
2         p05 = 0.05
3         zero = 0

          Dim tmpaarr() As Double
4         ReDim tmpaarr(1 To M)
          Dim tmpj As Single

          '
          '     epsmch is the machine precision.
          '
5         epsmch = 1.2E-16
          '
          '     compute the initial column norms and initialize several arrays.
          '
6         For j = 1 To N
7             For tmpj = 1 To M
8                 tmpaarr(tmpj) = a(tmpj, j)
9             Next tmpj
              'acnorm(j) = enorm(m, a(1, j))

              'a3 = enorm(m, a2)
              'testenorm tmpaarr
10            acnorm(j) = myenorm(M, tmpaarr)
11            rdiag(j) = acnorm(j)
12            wa(j) = rdiag(j)
13            If (pivot) Then ipvt(j) = j
14        Next j
          '
          '     reduce a to r with householder transformations.
          '
15        minmn = min1single(M, N)
16        For j = 1 To minmn
17            If (Not pivot) Then GoTo L40
              '
              '        bring the column of largest norm into the pivot position.
              '
18            kmax = j
19            For k = j To N
20                If (rdiag(k) > rdiag(kmax)) Then kmax = k
21            Next k
22            If (kmax = j) Then GoTo L40
23            For i = 1 To M
24                Temp = a(i, j)
25                a(i, j) = a(i, kmax)
26                a(i, kmax) = Temp
27            Next i
28            rdiag(kmax) = rdiag(j)
29            wa(kmax) = wa(j)
30            k = ipvt(j)
31            ipvt(j) = ipvt(kmax)
32            ipvt(kmax) = k
L40:
              '
              '        compute the householder transformation to reduce the
              '        j-th column of a to a multiple of the j-th unit vector.
              '
33            ReDim tmpaarr(1 To M - j + 1)
34            For tmpj = 1 To M - j + 1
35                tmpaarr(tmpj) = a(tmpj + j - 1, j)
36            Next tmpj
              'ajnorm = myenorm(m - j + 1, a(j, j))
37            ajnorm = myenorm(M - j + 1, tmpaarr)
38            If (ajnorm = zero) Then GoTo L100
39            If (a(j, j) < zero) Then ajnorm = -ajnorm
40            For i = j To M
41                a(i, j) = a(i, j) / ajnorm
42            Next i
43            a(j, j) = a(j, j) + one
              '
              '        apply the transformation to the remaining columns
              '        and update the norms.
              '
44            jp1 = j + 1
45            If (N < jp1) Then GoTo L100
46            For k = jp1 To N
47                sum = zero
48                For i = j To M
49                    sum = sum + a(i, j) * a(i, k)
50                Next i
51                Temp = sum / a(j, j)
52                For i = j To M
53                    a(i, k) = a(i, k) - Temp * a(i, j)
54                Next i
55                If ((Not pivot) Or (rdiag(k) = zero)) Then GoTo L80
56                Temp = a(j, k) / rdiag(k)
57                rdiag(k) = rdiag(k) * dsqrt(dmax1(zero, one - Temp * Temp))
58                If (p05 * (rdiag(k) / wa(k)) ^ 2 > epsmch) Then GoTo L80
59                rdiag(k) = myenorm(M - j, a(jp1, k))
60                wa(k) = rdiag(k)
L80:
61            Next k
L100:
62            rdiag(j) = -ajnorm
63        Next j
          'Return
          '
          '     last card of subroutine qrfac.
          '
End Sub
Private Function mymod(x As Single, y As Single) As Single
1         mymod = 1
End Function

Private Sub dogleg(ByRef N As Single, ByRef R As Variant, ByRef LR As Single, _
        ByRef diag As Variant, ByRef qtb As Variant, ByRef delta As Double, _
        ByRef x As Variant, ByRef WA1 As Variant, ByRef WA2 As Variant)
          'subroutine dogleg(n, r, lr, diag, qtb, delta, x, wa1, wa2)
          'integer n,lr
          'double precision delta
          'double precision r(lr),diag(n),qtb(n),x(n),wa1(n),wa2(n)
          Dim Alpha As Double
          Dim bnorm As Double
          Dim epsmch As Double
          Dim gnorm As Double
          Dim i As Single
          Dim j As Single
          Dim jj As Single
          Dim jp1 As Single
          Dim k As Single
          Dim L As Single
          Dim one As Double
          Dim qnorm As Double
          Dim sgnorm As Double
          Dim sum As Double
          Dim Temp As Double
          Dim zero As Double
1         one = 1
2         zero = 0
3         epsmch = 1E-16
          '      integer i,j,jj,jp1,k,l
          '      double precision alpha,bnorm,epsmch,gnorm,one,qnorm,sgnorm,sum,
          '     *                 temp,zero
          '      double precision dpmpar,enorm
          '      data one,zero /1.0d0,0.0d0/
          'epsmch = dpmpar(1)
          '
          '     first, calculate the gauss-newton direction.
          '
4         jj = (N * (N + 1)) / 2 + 1
5         For k = 1 To N
6             j = N - k + 1
7             jp1 = j + 1
8             jj = jj - k
9             L = jj + 1
10            sum = zero
11            If (N < jp1) Then GoTo L20
12            For i = jp1 To N
13                sum = sum + R(L) * x(i)
14                L = L + 1
15            Next i
L20:
16            Temp = R(jj)
17            If (Temp <> zero) Then GoTo L40
18            L = j
19            For i = 1 To j
20                Temp = dmax1(Temp, dabs(R(L)))
21                L = L + N - i
22            Next i
23            Temp = epsmch * Temp
24            If (Temp = zero) Then Temp = epsmch
L40:
25            x(j) = (qtb(j) - sum) / Temp
26        Next k
          '
          '     test whether the gauss-newton direction is acceptable.
          '
27        For j = 1 To N
28            WA1(j) = zero
29            WA2(j) = diag(j) * x(j)
30        Next j
31        qnorm = myenorm(N, WA2)
32        If (qnorm <= delta) Then GoTo L140
          '
          '     the gauss-newton direction is not acceptable.
          '     next, calculate the scaled gradient direction.
          '
33        L = 1
34        For j = 1 To N
35            Temp = qtb(j)
36            For i = j To N
37                WA1(i) = WA1(i) + R(L) * Temp
38                L = L + 1
39            Next i
40            WA1(j) = WA1(j) / diag(j)
41        Next j
          '
          '     calculate the norm of the scaled gradient and test for
          '     the special case in which the scaled gradient is zero.
          '
42        gnorm = myenorm(N, WA1)
43        sgnorm = zero
44        Alpha = delta / qnorm
45        If (gnorm = zero) Then GoTo L120
          '
          '     calculate the point along the scaled gradient
          '     at which the quadratic is minimized.
          '
46        For j = 1 To N
47            WA1(j) = (WA1(j) / gnorm) / diag(j)
48        Next j
49        L = 1
50        For j = 1 To N
51            sum = zero
52            For i = j To N
53                sum = sum + R(L) * WA1(i)
54                L = L + 1
55            Next i
56            WA2(j) = sum
57        Next j
58        Temp = myenorm(N, WA2)
59        sgnorm = (gnorm / Temp) / Temp
          '
          '     test whether the scaled gradient direction is acceptable.
          '
60        Alpha = zero
61        If (sgnorm >= delta) Then GoTo L120
          '
          '     the scaled gradient direction is not acceptable.
          '     finally, calculate the point along the dogleg
          '     at which the quadratic is minimized.
          '
62        bnorm = myenorm(N, qtb)
63        Temp = (bnorm / gnorm) * (bnorm / qnorm) * (sgnorm / delta)
64        Temp = Temp - (delta / qnorm) * (sgnorm / delta) ^ 2 _
              + dsqrt((Temp - (delta / qnorm)) ^ 2 _
              + (one - (delta / qnorm) ^ 2) * (one - (sgnorm / delta) ^ 2))
65        Alpha = ((delta / qnorm) * (one - (sgnorm / delta) ^ 2)) / Temp
L120:
          '
          '     form appropriate convex combination of the gauss-newton
          '     direction and the scaled gradient direction.
          '
66        Temp = (one - Alpha) * dmin1(sgnorm, delta)
67        For j = 1 To N
68            x(j) = Temp * WA1(j) + Alpha * x(j)
69        Next j
L140:
End Sub
Private Function dmin1(a As Double, b As Double)
1         If a <= b Then
2             dmin1 = a
3         Else
4             dmin1 = b
5         End If
End Function
Private Function dmax1(a, b)
1         If a >= b Then
2             dmax1 = a
3         Else
4             dmax1 = b
5         End If
End Function

Private Sub fdjac1(fcn2 As String, ByRef N As Single, ByRef x As Variant, ByRef fvec As Variant, _
        ByRef fjac As Variant, ByRef LDFJAC As Variant, ByRef iflag As Single, _
        ByRef ML As Single, ByRef mu As Single, ByRef EPSFCN As Double, _
        ByRef WA1() As Double, ByRef WA2 As Variant)

          ' Explicit double array WA1 PB#20150618

          'subroutine fdjac1(fcn,n,x,fvec,fjac,ldfjac,iflag,ml,mu,epsfcn,
          '    *                  wa1,wa2)
          '     integer n,ldfjac,iflag,ml,mu
          '     double precision epsfcn
          '     double precision x(n),fvec(n),fjac(ldfjac,n),wa1(n),wa2(n)
          Dim eps As Double
          Dim epsmch As Double
          Dim H As Double
          Dim i As Single
          Dim j As Single
          Dim k As Single
          Dim msum As Single
          Dim Temp As Double
          Dim zero As Double
          '          Dim dpmpar As Double
1         zero = 0

          '     integer i,j,k,msum
          '     double precision eps,epsmch,h,temp,zero
          '     double precision dpmpar
          '     data zero /0.0d0/
2         epsmch = 1E-16
          '
3         eps = dsqrt(dmax1(EPSFCN, epsmch))
4         msum = ML + mu + 1
5         If (msum < N) Then GoTo L40
          '
          '        computation of dense approximate jacobian.
          '
6         For j = 1 To N
7             Temp = x(j)
8             H = eps * dabs(Temp)
9             If (H = zero) Then H = eps
10            x(j) = Temp + H
              'redim paramArray(1 to
11            fcn N, x, WA1, iflag
12            If (iflag < 0) Then GoTo L30
13            x(j) = Temp
14            For i = 1 To N
15                fjac(i, j) = (WA1(i) - fvec(i)) / H
16            Next i
17        Next j
L30:
18        GoTo L110
L40:
          '
          '        computation of banded approximate jacobian.
          '
19        For k = 1 To msum
20            For j = k To N Step msum
21                WA2(j) = x(j)
22                H = eps * dabs(WA2(j))
23                If (H = zero) Then H = eps
24                x(j) = WA2(j) + H
25            Next j
26            fcn N, x, WA1, iflag
27            If (iflag < 0) Then GoTo L100
28            For j = k To N Step msum
29                x(j) = WA2(j)
30                H = eps * dabs(WA2(j))
31                If (H = zero) Then H = eps
32                For i = 1 To N
33                    fjac(i, j) = zero
34                    If ((i >= (j - mu)) And (i <= j + ML)) Then
35                        fjac(i, j) = (WA1(i) - fvec(i)) / H
36                    End If
37                Next i
38            Next j
39        Next k
L100:
L110:
End Sub

Public Function myenorm(N As Single, x As Variant) As Double
          Dim agiant As Double
          Dim floatn As Double
          Dim i As Single
          Dim one As Double
          Dim rdwarf As Double
          Dim rgiant As Double
          Dim s1 As Double
          Dim s2 As Double
          Dim s3 As Double
          Dim x1max As Double
          Dim x3max As Double
          Dim xabs As Double
          Dim zero As Double

1         one = 1
2         zero = 0
3         rdwarf = 3.834E-20
4         rgiant = 1.304E+19
5         s1 = zero
6         s2 = zero
7         s3 = zero
8         x1max = zero
9         x3max = zero
10        floatn = N
11        agiant = rgiant / floatn
12        For i = 1 To N
13            xabs = dabs(CDbl(x(i)))
14            If ((xabs > rdwarf) And (xabs < agiant)) Then GoTo L70
15            If (xabs <= rdwarf) Then GoTo L30
              '
              '              sum for large components.
              '
16            If (xabs <= x1max) Then GoTo L10
17            s1 = one + s1 * (x1max / xabs) ^ 2
18            x1max = xabs
19            GoTo L20
L10:
20            s1 = s1 + (xabs / x1max) ^ 2
L20:
21            GoTo L60
L30:
              '
              '              sum for small components.
              '
22            If (xabs <= x3max) Then GoTo L40
23            s3 = one + s3 * (x3max / xabs) ^ 2
24            x3max = xabs
25            GoTo L50
L40:
26            If (xabs <> zero) Then s3 = s3 + (xabs / x3max) ^ 2
L50:
L60:
27            GoTo L80
L70:
              '
              '           sum for intermediate components.
              '
28            s2 = s2 + xabs ^ 2
L80:
29        Next i
          '
          '     calculation of norm.
          '
30        If (s1 = zero) Then GoTo L100
31        myenorm = x1max * dsqrt(s1 + (s2 / x1max) / x1max)
32        GoTo L130
L100:
33        If (s2 = zero) Then GoTo L110
34        If (s2 >= x3max) Then _
              myenorm = dsqrt(s2 * (one + (x3max / s2) * (x3max * s3)))
35        If (s2 < x3max) Then _
              myenorm = dsqrt(x3max * ((s2 / x3max) + (x3max * s3)))
36        GoTo L120
L110:
37        myenorm = x3max * dsqrt(s3)
L120:
L130:
          '
          '     last card of function enorm.
          '
End Function

Private Function dsqrt(x)
1         dsqrt = x ^ 0.5
End Function
Private Function dabs(x)
1         dabs = Abs(x)
End Function
Private Sub r1mpyq(ByRef M As Single, ByRef N As Single, ByRef a As Variant, _
        ByRef lda As Single, ByRef v As Variant, ByRef w As Variant)

          'subroutine r1mpyq(m, n, a, lda, v, w)
          'integer m,n,lda
          'double precision a(lda,n),v(n),w(n)
          Dim i As Single
          Dim j As Single
          Dim mycos As Double
          Dim mysin As Double
          Dim nm1 As Single
          Dim nmj As Single
          Dim one As Double
          Dim Temp As Double
1         one = 1

          'integer i,j,nmj,nm1
          '      double precision cos,one,sin,temp
          '      data one /1.0d0/
          '
          '     apply the first set of givens rotations to a.
          '
2         nm1 = N - 1
3         If (nm1 < 1) Then GoTo L50
4         For nmj = 1 To nm1
5             j = N - nmj
6             If (dabs(v(j)) > one) Then mycos = one / v(j)
7             If (dabs(v(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
8             If (dabs(v(j)) <= one) Then mysin = v(j)
9             If (dabs(v(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
10            For i = 1 To M
11                Temp = mycos * a(i, j) - mysin * a(i, N)
12                a(i, N) = mysin * a(i, j) + mycos * a(i, N)
13                a(i, j) = Temp
14            Next i
15        Next nmj
          '
          '     apply the second set of givens rotations to a.
          '
16        For j = 1 To nm1
17            If (dabs(w(j)) > one) Then mycos = one / w(j)
18            If (dabs(w(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
19            If (dabs(w(j)) <= one) Then mysin = w(j)
20            If (dabs(w(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
21            For i = 1 To M
22                Temp = mycos * a(i, j) + mysin * a(i, N)
23                a(i, N) = -mysin * a(i, j) + mycos * a(i, N)
24                a(i, j) = Temp
25            Next i
26        Next j
L50:
End Sub

Private Sub r1mpyqforqtf(ByRef M As Single, ByRef N As Single, ByRef a As Variant, _
        ByRef lda As Single, ByRef v As Variant, ByRef w As Variant)

          'subroutine r1mpyq(m, n, a, lda, v, w)
          'integer m,n,lda
          'double precision a(lda,n),v(n),w(n)
          Dim i As Single
          Dim j As Single
          Dim mycos As Double
          Dim mysin As Double
          Dim nm1 As Single
          Dim nmj As Single
          Dim one As Double
          Dim Temp As Double
1         one = 1

          'integer i,j,nmj,nm1
          '      double precision cos,one,sin,temp
          '      data one /1.0d0/
          '
          '     apply the first set of givens rotations to a.
          '
2         nm1 = N - 1
3         If (nm1 < 1) Then GoTo L50
4         For nmj = 1 To nm1
5             j = N - nmj
6             If (dabs(v(j)) > one) Then mycos = one / v(j)
7             If (dabs(v(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
8             If (dabs(v(j)) <= one) Then mysin = v(j)
9             If (dabs(v(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
10            For i = 1 To M
11                Temp = mycos * a(j) - mysin * a(N)
12                a(N) = mysin * a(j) + mycos * a(N)
13                a(j) = Temp
14            Next i
15        Next nmj
          '
          '     apply the second set of givens rotations to a.
          '
16        For j = 1 To nm1
17            If (dabs(w(j)) > one) Then mycos = one / w(j)
18            If (dabs(w(j)) > one) Then mysin = dsqrt(one - mycos ^ 2)
19            If (dabs(w(j)) <= one) Then mysin = w(j)
20            If (dabs(w(j)) <= one) Then mycos = dsqrt(one - mysin ^ 2)
21            For i = 1 To M
22                Temp = mycos * a(j) + mysin * a(N)
23                a(N) = -mysin * a(j) + mycos * a(N)
24                a(j) = Temp
25            Next i
26        Next j
L50:
End Sub

Private Function min0(a, b)
1         If a <= b Then
2             min0 = a
3         Else
4             min0 = b
5         End If
End Function
Private Sub qform(ByRef M As Single, ByRef N As Single, ByRef q As Variant, _
        ByRef ldq As Single, ByRef wa As Variant)

          Dim i As Single
          Dim j As Single
          Dim jm1 As Single
          Dim k As Single
          Dim L As Single
          Dim minmn As Single
          Dim np1 As Single
          Dim one As Double
          Dim sum As Double
          Dim Temp As Double
          Dim zero As Double
1         one = 1
2         sum = 0
3         Temp = 1
4         zero = 0

          '     integer i,j,jm1,k,l,minmn,np1
          '    double precision one,sum,temp,zero
          '   data one,zero /1.0d0,0.0d0/
          '
          '     zero out upper triangle of q in the first min(m,n) columns.
          '
5         minmn = min0(M, N)
6         If minmn < 2 Then GoTo L30
7         For j = 2 To minmn
8             jm1 = j - 1
9             For i = 1 To jm1
10                q(i, j) = zero
11            Next i
12        Next j

L30:
          '
          '     initialize remaining columns to those of the identity matrix.
          '
13        np1 = N + 1
14        If (M < np1) Then GoTo L60
15        For j = np1 To M
16            For i = 1 To M
17                q(i, j) = zero
18            Next i
19            q(j, j) = one
20        Next j
L60:
          '
          '     accumulate q from its factored form.
          '
21        For L = 1 To minmn
22            k = minmn - L + 1
23            For i = k To M
24                wa(i) = q(i, k)
25                q(i, k) = zero
26            Next i
27            q(k, k) = one
28            If (wa(k) = zero) Then GoTo L110
29            For j = k To M
30                sum = zero
31                For i = k To M
32                    sum = sum + q(i, j) * wa(i)
33                Next i
34                Temp = sum / wa(k)
35                For i = k To M
36                    q(i, j) = q(i, j) - Temp * wa(i)
37                Next i
38            Next j
L110:
39        Next L
End Sub
