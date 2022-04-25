Attribute VB_Name = "modBroadcast"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modBroadcast
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : Functions to create array processing functions from singleton-processing functions
'             Is there a better way to do this using class interfaces? - would be easy in functional programming language...
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module        'We don't want to expose these functions to Excel
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Broadcast1Arg
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Version of Broadcast for methods that take only a single argument. If Arg1 is not an array then return is not an array
'             if Arg1 is an array then return is a 2-dimensional array (even if arg1 is 1 dimensional)
' -----------------------------------------------------------------------------------------------------------------------
Function Broadcast1Arg(CoreFunctionID As BroadcastFuncID, ByVal Arg1 As Variant)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result()

1         On Error GoTo ErrHandler
2         If VarType(Arg1) < vbArray Then
              'When adding cases, also edit the second Select Case statement
3             Select Case CoreFunctionID
                  Case FuncIdFileDelete
4                     Broadcast1Arg = CoreFileDelete(CStr(Arg1))
5                 Case FuncIdFileExists
6                     Broadcast1Arg = CoreFileExists(CStr(Arg1))
7                 Case FuncIdFileLastModifiedDate
8                     Broadcast1Arg = CoreFileLastModifiedDate(CStr(Arg1))
9                 Case FuncIdFileNumLines
10                    Broadcast1Arg = CoreFileNumLines(CStr(Arg1))
11                Case FuncIdCreateFolder
12                    Broadcast1Arg = CoreCreateFolder(CStr(Arg1))
13                Case FuncIdDeleteFolder
14                    Broadcast1Arg = CoreDeleteFolder(CStr(Arg1))
15                Case FuncIdFolderExists
16                    Broadcast1Arg = CoreFolderExists(CStr(Arg1))
17                Case FuncIdFolderIsWritable
18                    Broadcast1Arg = CoreFolderIsWritable(CStr(Arg1))
19                Case FuncIdFileIsUnicode
20                    Broadcast1Arg = IsUnicodeFile(CStr(Arg1))
21                Case FuncIdFileUnblock
22                    Broadcast1Arg = CoreFileUnblock(CStr(Arg1))
23            End Select
24        Else
25            Force2DArrayR Arg1, NR, NC
26            ReDim Result(1 To NR, 1 To NC)
              'When adding cases, also edit the first Select Case statement
27            Select Case CoreFunctionID
                  Case FuncIdFileDelete
28                    For i = 1 To NR
29                        For j = 1 To NC
30                            Result(i, j) = CoreFileDelete(CStr(Arg1(i, j)))
31                        Next j
32                    Next i
33                Case FuncIdFileExists
34                    For i = 1 To NR
35                        For j = 1 To NC
36                            Result(i, j) = CoreFileExists(CStr(Arg1(i, j)))
37                        Next j
38                    Next i
39                Case FuncIdFileLastModifiedDate
40                    For i = 1 To NR
41                        For j = 1 To NC
42                            Result(i, j) = CoreFileLastModifiedDate(CStr(Arg1(i, j)))
43                        Next j
44                    Next i
45                Case FuncIdFileNumLines
46                    For i = 1 To NR
47                        For j = 1 To NC
48                            Result(i, j) = CoreFileNumLines(CStr(Arg1(i, j)))
49                        Next j
50                    Next i
51                Case FuncIdCreateFolder
52                    For i = 1 To NR
53                        For j = 1 To NC
54                            Result(i, j) = CoreCreateFolder(CStr(Arg1(i, j)))
55                        Next j
56                    Next i
57                Case FuncIdDeleteFolder
58                    For i = 1 To NR
59                        For j = 1 To NC
60                            Result(i, j) = CoreDeleteFolder(CStr(Arg1(i, j)))
61                        Next j
62                    Next i
63                Case FuncIdFolderExists
64                    For i = 1 To NR
65                        For j = 1 To NC
66                            Result(i, j) = CoreFolderExists(CStr(Arg1(i, j)))
67                        Next j
68                    Next i
69                Case FuncIdFolderIsWritable
70                    For i = 1 To NR
71                        For j = 1 To NC
72                            Result(i, j) = CoreFolderIsWritable(CStr(Arg1(i, j)))
73                        Next j
74                    Next i
75                Case FuncIdFileIsUnicode
76                    For i = 1 To NR
77                        For j = 1 To NC
78                            Result(i, j) = IsUnicodeFile(CStr(Arg1(i, j)))
79                        Next j
80                    Next i
81                Case FuncIdFileUnblock
82                    For i = 1 To NR
83                        For j = 1 To NC
84                            Result(i, j) = CoreFileUnblock(CStr(Arg1(i, j)))
85                        Next j
86                    Next i
87            End Select
88            Broadcast1Arg = Result
89        End If
90        Exit Function
ErrHandler:
91        Broadcast1Arg = "#Broadcast1Arg (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Broadcast2Args
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Version of Broadcast where we have two args and both must have the same number of rows and cols
' -----------------------------------------------------------------------------------------------------------------------
Function Broadcast2Args(CoreFunctionID As BroadcastFuncID, ByVal Arg1 As Variant, ByVal Arg2 As Variant, Optional ByVal Arg3 As Variant, Optional ByVal Arg4 As Variant, Optional ByVal Arg5 As Variant)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result()

1         On Error GoTo ErrHandler
2         If VarType(Arg1) < vbArray And VarType(Arg2) < vbArray Then
              'When adding cases, also edit the second Select Case statement
3             Select Case CoreFunctionID
                  Case FuncIdFileCopy
4                     Broadcast2Args = CoreFileCopy(CStr(Arg1), CStr(Arg2), CBool(Arg3))
5                 Case FuncIdFileMove
6                     Broadcast2Args = CoreFileMove(CStr(Arg1), CStr(Arg2))
7                 Case FuncIdFileRename
8                     Broadcast2Args = CoreFileRename(CStr(Arg1), CStr(Arg2))
9                 Case FuncIdFolderRename
10                    Broadcast2Args = CoreFolderRename(CStr(Arg1), CStr(Arg2))
11                Case FuncIdURLDownloadToFile
12                    Broadcast2Args = CoreURLDownloadToFile(CStr(Arg1), CStr(Arg2))
13                Case FuncIdFileRegExReplace
14                    Broadcast2Args = CoreFileRegExReplace(CStr(Arg1), CStr(Arg2), CStr(Arg3), CStr(Arg4), CBool(Arg5))
15                Case FuncIdFolderCopy
16                    Broadcast2Args = CoreFolderCopy(CStr(Arg1), CStr(Arg2))
17                Case FuncIdFolderMove
18                    Broadcast2Args = CoreFolderMove(CStr(Arg1), CStr(Arg2))
19                Case FuncIDFileTranspose
20                    Broadcast2Args = CoreFileTranspose(CStr(Arg1), CStr(Arg2), CStr(Arg3))
21            End Select
22        Else
23            Force2DArrayRMulti Arg1, Arg2
24            NR = sNRows(Arg1)
25            NC = sNCols(Arg1)
26            If sNRows(Arg2) <> NR Then Throw "First two arguments must be arrays of the same size"
27            If sNCols(Arg2) <> NC Then Throw "First two arguments must be arrays of the same size"
28            ReDim Result(1 To NR, 1 To NC)
              'When adding cases, also edit the first Select Case statement
29            Select Case CoreFunctionID
                  Case FuncIdFileCopy
30                    For i = 1 To NR
31                        For j = 1 To NC
32                            Result(i, j) = CoreFileCopy(CStr(Arg1(i, j)), CStr(Arg2(i, j)), CBool(Arg3))
33                        Next j
34                    Next i
35                Case FuncIdFileMove
36                    For i = 1 To NR
37                        For j = 1 To NC
38                            Result(i, j) = CoreFileMove(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
39                        Next j
40                    Next i
41                Case FuncIdFileRename
42                    For i = 1 To NR
43                        For j = 1 To NC
44                            Result(i, j) = CoreFileRename(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
45                        Next j
46                    Next i
47                Case FuncIdFolderRename
48                    For i = 1 To NR
49                        For j = 1 To NC
50                            Result(i, j) = CoreFolderRename(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
51                        Next j
52                    Next i
53                Case FuncIdURLDownloadToFile
54                    For i = 1 To NR
55                        For j = 1 To NC
56                            Result(i, j) = CoreURLDownloadToFile(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
57                        Next j
58                    Next i
59                Case FuncIdFileRegExReplace
60                    For i = 1 To NR
61                        For j = 1 To NC
62                            Result(i, j) = CoreFileRegExReplace(CStr(Arg1(i, j)), CStr(Arg2(i, j)), CStr(Arg3), CStr(Arg4), CBool(Arg5))
63                        Next j
64                    Next i
65                Case FuncIdFolderCopy
66                    For i = 1 To NR
67                        For j = 1 To NC
68                            Result(i, j) = CoreFolderCopy(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
69                        Next j
70                    Next i
71                Case FuncIdFolderMove
72                    For i = 1 To NR
73                        For j = 1 To NC
74                            Result(i, j) = CoreFolderMove(CStr(Arg1(i, j)), CStr(Arg2(i, j)))
75                        Next j
76                    Next i
77                Case FuncIDFileTranspose
78                    For i = 1 To NR
79                        For j = 1 To NC
80                            Result(i, j) = CoreFileTranspose(CStr(Arg1(i, j)), CStr(Arg2(i, j)), CStr(Arg3))
81                        Next j
82                    Next i
83            End Select
84            Broadcast2Args = Result
85        End If
86        Exit Function
ErrHandler:
87        Broadcast2Args = "#Broadcast2Args (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Broadcast
' Author    : Philip Swannell
' Date      : 04-May-2015
' Purpose   : Transform a function all of whose arguments are singletons into a method that does
'             element-wise array processing.
'             16 May 15 note: First versions of this method used ParamArray, but ParamArray arguments
'             can only be passed by reference and since we need to force the arguments to be 2-dimensional
'             that meant I needed a complicated scheme of backing up and restoring arguments so that the
'             calling method did not see changes to them. In the end it was simpler to stop using
'             ParamArray at the cost of repeated blocks of code at the top of the method.
' -----------------------------------------------------------------------------------------------------------------------
Function Broadcast(CoreFunctionID As BroadcastFuncID, ByVal Arg1 As Variant, Optional ByVal Arg2 As Variant, Optional ByVal Arg3 As Variant, _
          Optional ByVal Arg4 As Variant, Optional ByVal Arg5 As Variant, Optional ByVal Arg6 As Variant, Optional ByVal Arg7 As Variant, _
          Optional ByVal Arg8 As Variant, Optional ByVal Arg9 As Variant, Optional ByVal Arg10 As Variant)

          Dim AnyArgs As Boolean
          Dim c As Long
          Dim ColLock(1 To 10) As Boolean
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim R As Long
          Dim Result()
          Dim RowLock(1 To 10) As Boolean
          Dim MissingArray(1 To 1, 1 To 1) As Variant

1         On Error GoTo ErrHandler

2         MissingArray(1, 1) = CreateMissing()

3         NR = 1: NC = 1
4         If Not (IsMissing(Arg1)) Then
5             Force2DArrayR Arg1, R, c
6             AnyArgs = True
7             If R = 1 Then RowLock(1) = True
8             If c = 1 Then ColLock(1) = True
9             If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
10            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
11        End If
12        If Not (IsMissing(Arg2)) Then
13            Force2DArrayR Arg2, R, c
14            AnyArgs = True
15            If R = 1 Then RowLock(2) = True
16            If c = 1 Then ColLock(2) = True
17            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
18            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
19        End If
20        If Not (IsMissing(Arg3)) Then
21            Force2DArrayR Arg3, R, c
22            AnyArgs = True
23            If R = 1 Then RowLock(3) = True
24            If c = 1 Then ColLock(3) = True
25            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
26            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
27        End If
28        If Not (IsMissing(Arg4)) Then
29            Force2DArrayR Arg4, R, c
30            AnyArgs = True
31            If R = 1 Then RowLock(4) = True
32            If c = 1 Then ColLock(4) = True
33            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
34            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
35        End If
36        If Not (IsMissing(Arg5)) Then
37            Force2DArrayR Arg5, R, c
38            AnyArgs = True
39            If R = 1 Then RowLock(5) = True
40            If c = 1 Then ColLock(5) = True
41            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
42            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
43        End If
44        If Not (IsMissing(Arg6)) Then
45            Force2DArrayR Arg6, R, c
46            AnyArgs = True
47            If R = 1 Then RowLock(6) = True
48            If c = 1 Then ColLock(6) = True
49            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
50            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
51        End If
52        If Not (IsMissing(Arg7)) Then
53            Force2DArrayR Arg7, R, c
54            AnyArgs = True
55            If R = 1 Then RowLock(7) = True
56            If c = 1 Then ColLock(7) = True
57            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
58            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
59        End If
60        If Not (IsMissing(Arg8)) Then
61            Force2DArrayR Arg8, R, c
62            AnyArgs = True
63            If R = 1 Then RowLock(8) = True
64            If c = 1 Then ColLock(8) = True
65            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
66            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
67        End If
68        If Not (IsMissing(Arg9)) Then
69            Force2DArrayR Arg9, R, c
70            AnyArgs = True
71            If R = 1 Then RowLock(9) = True
72            If c = 1 Then ColLock(9) = True
73            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
74            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
75        End If
76        If Not (IsMissing(Arg10)) Then
77            Force2DArrayR Arg10, R, c
78            AnyArgs = True
79            If R = 1 Then RowLock(10) = True
80            If c = 1 Then ColLock(10) = True
81            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
82            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
83        End If

84        If Not AnyArgs Then Throw "No arguments supplied"

85        ReDim Result(1 To NR, 1 To NC)

86        Select Case CoreFunctionID
              Case FuncIdBlackScholes
87                For i = 1 To NR
88                    For j = 1 To NC
89                        Result(i, j) = bsCore(StringToOptStyle(Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)), True), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              (Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))))
90                    Next j
91                Next i
92            Case FuncIdNormOpt
93                For i = 1 To NR
94                    For j = 1 To NC
95                        Result(i, j) = CoreNormOpt(StringToOptStyle(Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)), True), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              (Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))))
96                    Next j
97                Next i
98            Case FuncIdOptSolveVol
99                For i = 1 To NR
100                   For j = 1 To NC
101                       Result(i, j) = CoreOptSolveVol(StringToOptStyle(Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)), False), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              (Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))), _
                              (Arg6(IIf(RowLock(6), 1, i), IIf(ColLock(6), 1, j))))
102                   Next j
103               Next i
104           Case FuncIdDivide
105               For i = 1 To NR
106                   For j = 1 To NC
107                       Result(i, j) = SafeDivide((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
108                   Next j
109               Next i
110           Case FuncIdSubtract
111               For i = 1 To NR
112                   For j = 1 To NC
113                       Result(i, j) = SafeSubtract((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
114                   Next j
115               Next i
116           Case FuncIdEquals
117               For i = 1 To NR
118                   For j = 1 To NC
119                       Result(i, j) = sEquals((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))))
120                   Next j
121               Next i
122           Case FuncIdPower
123               For i = 1 To NR
124                   For j = 1 To NC
125                       Result(i, j) = SafePower((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
126                   Next j
127               Next i
128           Case FuncIdStringBetweenStrings
129               For i = 1 To NR
130                   For j = 1 To NC
131                       Result(i, j) = CoreStringBetweenStrings((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              (Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))))
132                   Next j
133               Next i
134           Case FuncIdLessThan
135               For i = 1 To NR
136                   For j = 1 To NC
137                       Result(i, j) = VariantLessThan((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), True)
138                   Next j
139               Next i
140           Case FuncIdlessThanOrEqual
141               For i = 1 To NR
142                   For j = 1 To NC
143                       Result(i, j) = VariantLessThanOrEqual((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), True)
144                   Next j
145               Next i
146           Case FuncIdIf
147               For i = 1 To NR
148                   For j = 1 To NC
149                       Result(i, j) = SafeIf((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))))
150                   Next j
151               Next i
152           Case FuncIdLeft
153               For i = 1 To NR
154                   For j = 1 To NC
155                       Result(i, j) = SafeLeft((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
156                   Next j
157               Next i
158           Case FuncIdRight
159               For i = 1 To NR
160                   For j = 1 To NC
161                       Result(i, j) = SafeRight((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
162                   Next j
163               Next i
164           Case FuncIdNearlyEquals
165               For i = 1 To NR
166                   For j = 1 To NC
167                       Result(i, j) = sNearlyEquals((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))))
168                   Next j
169               Next i
170           Case FuncIdIfErrorString
171               For i = 1 To NR
172                   For j = 1 To NC
173                       Result(i, j) = SafeIfErrorString((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
174                   Next j
175               Next i
176           Case FuncIdLike
177               For i = 1 To NR
178                   For j = 1 To NC
179                       Result(i, j) = SafeLike(CStr((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)))), _
                              CStr((Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j)))))
180                   Next j
181               Next i
182           Case FuncIdRoundSF
183               For i = 1 To NR
184                   For j = 1 To NC
185                       Result(i, j) = CoreRoundSF((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))))
186                   Next j
187               Next i
188           Case FuncIDRound
189               For i = 1 To NR
190                   For j = 1 To NC
191                       Result(i, j) = CoreRound((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))))
192                   Next j
193               Next i
194           Case FuncIdFileInfo
195               For i = 1 To NR
196                   For j = 1 To NC
197                       Result(i, j) = CoreFileInfo((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
198                   Next j
199               Next i
200           Case FuncIdFileExif
201               For i = 1 To NR
202                   For j = 1 To NC
203                       Result(i, j) = CoreFileExif((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
204                   Next j
205               Next i
206           Case FuncIdFileCopySkip
207               For i = 1 To NR
208                   For j = 1 To NC
209                       Result(i, j) = CoreFileCopySkip((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))))
210                   Next j
211               Next i
212           Case FuncIDEDate
213               For i = 1 To NR
214                   For j = 1 To NC
215                       Result(i, j) = CoreEDate((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
216                   Next j
217               Next i
218           Case FuncIdBarrierOption
219               For i = 1 To NR
220                   For j = 1 To NC
221                       Result(i, j) = BarrierOption((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              (Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              (Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              (Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))), _
                              (Arg6(IIf(RowLock(6), 1, i), IIf(ColLock(6), 1, j))), _
                              (Arg7(IIf(RowLock(7), 1, i), IIf(ColLock(7), 1, j))), _
                              (Arg8(IIf(RowLock(8), 1, i), IIf(ColLock(8), 1, j))), _
                              (Arg9(IIf(RowLock(9), 1, i), IIf(ColLock(9), 1, j))), _
                              (Arg10(IIf(RowLock(10), 1, i), IIf(ColLock(10), 1, j))))
222                   Next j
223               Next i
224           Case FuncIDISDASIMMMakeID
225               For i = 1 To NR
226                   For j = 1 To NC
227                       Result(i, j) = ISDASIMMMakeID_Core(CStr((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)))), _
                              CStr(Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))), _
                              CStr(Arg3(IIf(RowLock(3), 1, i), IIf(ColLock(3), 1, j))), _
                              CStr(Arg4(IIf(RowLock(4), 1, i), IIf(ColLock(4), 1, j))), _
                              CStr(Arg5(IIf(RowLock(5), 1, i), IIf(ColLock(5), 1, j))), _
                              CLng(Arg6(IIf(RowLock(6), 1, i), IIf(ColLock(6), 1, j))))
228                   Next j
229               Next i
230           Case FuncIdISDASIMMApplyRounding2022
231               For i = 1 To NR
232                   For j = 1 To NC
233                       Result(i, j) = ISDASIMMApplyRounding2022((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
234                   Next j
235               Next i
236           Case FuncIdRelativePath
237               For i = 1 To NR
238                   For j = 1 To NC
239                       Result(i, j) = CoreRelativePath((Arg1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                              (Arg2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
240                   Next j
241               Next i
242           Case Else
243               Throw "Unrecognised CoreFunctionID"
244       End Select
245       Broadcast = Result
246       Exit Function
ErrHandler:
247       Broadcast = "#Broadcast (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'============================================================================== END OF BROADCAST FUNCTIONS
'BROADCASTASSOCIATIVE FUNCTIONS
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BroadcastAssociative
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Broadcast to handle functions that take an arbitrary number of arguments, and where
'             f(a,b,c,d) = f(f(f(a,b),c),d) i.e. function is associative.
' -----------------------------------------------------------------------------------------------------------------------
Function BroadcastAssociative(CoreFunctionID As BroadcastFuncID, ParamArray TheArguments())
          Dim Arguments()
          Dim c As Long
          Dim ColLock() As Boolean
          Dim FNM As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim LB As Long
          Dim NC As Long
          Dim NR As Long
          Dim R As Long
          Dim Result()
          Dim RowLock() As Boolean
          Dim UB As Long

1         On Error GoTo ErrHandler
2         Arguments = TheArguments(0)        'This line looks strange. It's because the wrapper function (e.g. sArrayAdd) _
                                              is itself declared with a ParamArray argument: Function sArrayAdd(ParamArray ArraysToAdd()) so ArraysToAdd is 1 one _
                                              dimensional array of (typically) 2-d arrays. By the time ArraysToAdd is passed to BroadcastAssociative, ArraysToAdd _
                                              becomes the zeroth element of the one-dimensional array TheArguments. So we need Arguments = TheArguments(0) to _
                                              extract the array of arguments to sArrayAdd - which is what we want to iterate through. _
                                              The line Arguments = TheArguments(0) also has the advantage that changes made by calls to Force2DArrayR are _
                                              invisible to  callers of sArrayAdd.

3         NR = 1: NC = 1
4         LB = LBound(Arguments)
5         UB = UBound(Arguments)
6         If UB < LB Then Throw "No arguments supplied"
7         ReDim RowLock(LB To UB)
8         ReDim ColLock(LB To UB)

9         For i = LB To UB
10            If Not IsMissing(Arguments(i)) Then
11                Force2DArrayR Arguments(i), R, c
12                If R = 1 Then RowLock(i) = True
13                If c = 1 Then ColLock(i) = True
14                If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
15                If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
16            End If
17        Next i

18        ReDim Result(1 To NR, 1 To NC)

19        FNM = -2
20        For k = LB To UB
21            If Not IsMissing(Arguments(k)) Then
22                FNM = k        'FNM = FirstNotMissing
23                Exit For
24            End If
25        Next k
26        If FNM = -2 Then Throw "All input arrays are missing"

27        For i = 1 To NR
28            For j = 1 To NC
29                Result(i, j) = Arguments(FNM)(IIf(RowLock(FNM), 1, i), IIf(ColLock(FNM), 1, j))
30            Next j
31        Next i

32        For k = FNM + 1 To UBound(Arguments)
33            If Not IsMissing(Arguments(k)) Then
34                For i = 1 To NR
35                    For j = 1 To NC
36                        If CoreFunctionID = FuncIdMultiply Then
37                            Result(i, j) = SafeMultiply(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
38                        ElseIf CoreFunctionID = FuncIdAdd Then
39                            Result(i, j) = SafeAdd(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
40                        ElseIf CoreFunctionID = FuncIdMax Then
41                            Result(i, j) = SafeMax(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
42                        ElseIf CoreFunctionID = FuncIdMin Then
43                            Result(i, j) = SafeMin(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
44                        ElseIf CoreFunctionID = FuncIdConcatenate Then
45                            Result(i, j) = SafeConcatenate(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
46                        ElseIf CoreFunctionID = FuncIdAnd Then
47                            Result(i, j) = SafeAnd(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
48                        ElseIf CoreFunctionID = FuncIdor Then
49                            Result(i, j) = SafeOr(Result(i, j), Arguments(k)(IIf(RowLock(k), 1, i), IIf(ColLock(k), 1, j)))
50                        Else
51                            Throw "Unrecognised CoreFunctionID"
52                        End If
53                    Next j
54                Next i
55            End If
56        Next k

57        BroadcastAssociative = Result
58        Exit Function
ErrHandler:
59        BroadcastAssociative = "#BroadcastAssociative (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BroadcastColumn
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Broadcast function to implement sColumnAnd, sColumnMax, sColumnMean, sColumnMin, sColumnOr, sColumnProduct, sColumnStDev, sColumnSum
' -----------------------------------------------------------------------------------------------------------------------
Function BroadcastColumn(CoreFunctionID As BroadcastFuncID, TheArray As Variant)
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         ReDim Result(1 To 1, 1 To M)

4         Select Case CoreFunctionID
              Case FuncIdAdd
5                 For j = 1 To M
6                     Result(1, j) = TheArray(1, j)
7                     For i = 2 To N
8                         Result(1, j) = SafeAdd(Result(1, j), TheArray(i, j))
9                     Next i
10                Next j
11            Case FuncIdMax
12                For j = 1 To M
13                    Result(1, j) = TheArray(1, j)
14                    For i = 2 To N
15                        Result(1, j) = SafeMax(Result(1, j), TheArray(i, j))
16                    Next i
17                Next j
18            Case FuncIdMin
19                For j = 1 To M
20                    Result(1, j) = TheArray(1, j)
21                    For i = 2 To N
22                        Result(1, j) = SafeMin(Result(1, j), TheArray(i, j))
23                    Next i
24                Next j
25            Case FuncIdMultiply
26                For j = 1 To M
27                    Result(1, j) = TheArray(1, j)
28                    For i = 2 To N
29                        Result(1, j) = SafeMultiply(Result(1, j), TheArray(i, j))
30                    Next i
31                Next j
32            Case FuncIdAnd
33                For j = 1 To M
34                    Result(1, j) = TheArray(1, j)
35                    For i = 2 To N
36                        Result(1, j) = SafeAnd(Result(1, j), TheArray(i, j))
37                    Next i
38                Next j
39            Case FuncIdor
40                For j = 1 To M
41                    Result(1, j) = TheArray(1, j)
42                    For i = 2 To N
43                        Result(1, j) = SafeOr(Result(1, j), TheArray(i, j))
44                    Next i
45                Next j
46        End Select

47        BroadcastColumn = Result
48        Exit Function
ErrHandler:
49        BroadcastColumn = "#BroadcastColumn (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BroadcastRow
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Broadcast function to implement sRowAnd, sRowMax, sRowMean, sRowMin, sRowOr, sRowProduct, sRowStDev, sRowSum
' -----------------------------------------------------------------------------------------------------------------------
Function BroadcastRow(CoreFunctionID As BroadcastFuncID, TheArray As Variant)
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         ReDim Result(1 To N, 1 To 1)

4         Select Case CoreFunctionID
              Case FuncIdAdd
5                 For i = 1 To N
6                     Result(i, 1) = TheArray(i, 1)
7                     For j = 2 To M
8                         Result(i, 1) = SafeAdd(Result(i, 1), TheArray(i, j))
9                     Next j
10                Next i
11            Case FuncIdMax
12                For i = 1 To N
13                    Result(i, 1) = TheArray(i, 1)
14                    For j = 2 To M
15                        Result(i, 1) = SafeMax(Result(i, 1), TheArray(i, j))
16                    Next j
17                Next i
18            Case FuncIdMin
19                For i = 1 To N
20                    Result(i, 1) = TheArray(i, 1)
21                    For j = 2 To M
22                        Result(i, 1) = SafeMin(Result(i, 1), TheArray(i, j))
23                    Next j
24                Next i
25            Case FuncIdMultiply
26                For i = 1 To N
27                    Result(i, 1) = TheArray(i, 1)
28                    For j = 2 To M
29                        Result(i, 1) = SafeMultiply(Result(i, 1), TheArray(i, j))
30                    Next j
31                Next i
32            Case FuncIdAnd
33                For i = 1 To N
34                    Result(i, 1) = TheArray(i, 1)
35                    For j = 2 To M
36                        Result(i, 1) = SafeAnd(Result(i, 1), TheArray(i, j))
37                    Next j
38                Next i
39            Case FuncIdor
40                For i = 1 To N
41                    Result(i, 1) = TheArray(i, 1)
42                    For j = 2 To M
43                        Result(i, 1) = SafeOr(Result(i, 1), TheArray(i, j))
44                    Next j
45                Next i
46        End Select

47        BroadcastRow = Result
48        Exit Function
ErrHandler:
49        BroadcastRow = "#BroadcastRow (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
