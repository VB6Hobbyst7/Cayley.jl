VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSobol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Changwei Xiong (axcw@hotmail.com), March 2013, http://www.cs.utah.edu/~cxiong/
' For more details about the algorithm, please refer to the link below:
'   http://web.maths.unsw.edu.au/~fkuo/sobol/index.html
' This VBA code is essentially translated from their C++ code
'Changes by Philip Swannell 7 Jun 2015
'1) Added error handlers to all methods for easy debugging
'2) Change in method SetData to cope with 1 dimensional case

Const MaxBit As Long = 31        'Not 32 because VBA doesn't support unsigned long

Private D As Long        ' dimension
Private N As Long        ' total # points in the sequence
Private with_shift As Boolean

Private H As Long        ' # of bits for n

Private v() As Long        ' v vectors
Private x() As Long        ' the current point (Integer) in the sequence
Private y() As Double        ' the current point (Decimal) in the sequence

Public Sub SetData(d_ As Long, n_ As Long, with_shift_ As Boolean)
1         On Error GoTo ErrHandler
2         D = d_        ' dimension
3         N = n_        ' total # points in sequence
4         with_shift = with_shift_
5         H = Application.WorksheetFunction.Ceiling(Log(N + 1) / Log(2), 1)

6         ReDim x(1 To D) As Long        ' the initial point starts from zero
7         ReDim y(1 To N, 1 To D) As Double        ' array for sobol sequence

8         If D = 1 Then        ' this block added by Philip Swannell 7 June 2015
9             y = sGrid(1 / (N + 1), 1 - 1 / (N + 1), N)
10            If with_shift_ Then
                  Dim i As Long
11                For i = 1 To N
12                    y(i, 1) = y(i, 1) + Rnd
13                Next i
14            End If
15            Exit Sub
16        End If

17        InitVectorV
18        SetSobol
19        Exit Sub
ErrHandler:
20        Throw "#SetData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Public Function GetSobolSequence() As Double()
1         On Error GoTo ErrHandler
2         GetSobolSequence = y
3         Exit Function
ErrHandler:
4         Throw "#GetSobolSequence (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub SetSobol()
          ' generate sobol sequence based on the v vector
          Dim i As Long
          Dim j As Long
          Dim u() As Double
1         On Error GoTo ErrHandler
2         ReDim u(1 To D) As Double        ' random shift
3         For j = 1 To D
4             x(j) = 0
5             If with_shift Then u(j) = Rnd Else u(j) = 0
6         Next j

          Dim F As Double
7         ReDim y(1 To N, 1 To D) As Double        ' array for sobol sequence
8         For i = 1 To N
9             Dim z As Long: z = FirstZeroBitIndex(i - 1)        'index MUST start from zero
10            For j = 1 To D
11                x(j) = x(j) Xor v(z, j)        ' x(i+1,j) = x(i-1,j) (+) v(z,j)
12                F = x(j) / 2 ^ MaxBit + u(j)        ' convert to a decimal between (0,1)
13                If F >= 1 Then F = F - 1
14                If F = 0 Then F = 1E-16        ' prevent it from being zero
15                y(i, j) = F
16            Next j
17        Next i
18        Exit Sub
ErrHandler:
19        Throw "#SetSobol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub InitVectorV()
          ' fetch the primative polynomial data
          Dim avec() As Long
          Dim M() As Long
          Dim poly As Variant
          Dim pvec() As Long
1         On Error GoTo ErrHandler
2         poly = shSobolPolynomial.Range("TopLeftCell").Resize(D, 21)        ' Remember to UPDATE this if sheet name changes
3         ReDim pvec(2 To D) As Long
4         ReDim avec(2 To D) As Long
5         ReDim M(1 To 21, 2 To D) As Long
          Dim j As Long
          Dim k As Long
          Dim q As Long
6         For j = 2 To D
7             pvec(j) = poly(j, 2)
8             avec(j) = poly(j, 3)
9             For k = 1 To pvec(j)
10                M(k, j) = poly(j, k + 3)
11            Next k
12        Next j

          ' prepare for v vector
13        ReDim v(1 To H, 1 To D) As Long
          ' the 1st dimension
14        For k = 1 To H        ' the 1st dimensin assumes all m_i'p = 1
15            v(k, 1) = 1 * 2 ^ (MaxBit - k)
16        Next k
          ' the higher dimensions
          Dim a As Long
          Dim p As Long
          Dim vk As Long
17        For j = 2 To D
18            p = pvec(j)
19            a = avec(j)
20            For k = 1 To H
21                If k <= p Then        ' m_i for 1<i<=p are given
22                    v(k, j) = M(k, j) * 2 ^ (MaxBit - k)
23                Else        ' recursive definition
24                    vk = v(k - p, j) Xor (v(k - p, j) \ 2 ^ p)        ' ~= v(k-p) (+) v(k-p)\2^p
25                    For q = 1 To p - 1        ' = XOR_Sum of a(q)*v(k-q) for 1<=q<=p-1
26                        vk = vk Xor (((a \ 2 ^ (p - 1 - q)) And 1) * v(k - q, j))
27                    Next q        ' binary a = (a[1],a[2],...,a[p-1])
28                    v(k, j) = vk
29                End If
30            Next k
31        Next j
32        Exit Sub
ErrHandler:
33        Throw "#InitVectorV (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function FirstZeroBitIndex(ByVal index As Long) As Long
          ' return the position of the first zero bit of an Index number (integer)
1         Dim z As Long: z = 1
2         On Error GoTo ErrHandler
3         Do While index And 1
4             z = z + 1
5             index = index \ 2
6         Loop
7         FirstZeroBitIndex = z
8         Exit Function
ErrHandler:
9         Throw "#FirstZeroBitIndex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
