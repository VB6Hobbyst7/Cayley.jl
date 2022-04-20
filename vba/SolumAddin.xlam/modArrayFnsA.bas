Attribute VB_Name = "modArrayFnsA"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modArrayFunction
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Functions for manipulating arrays of arbitrary data. Functions implemented
'             using the Broadcast... suite of functions go in ModArrayFunctionsBroadcast.
'             This module is for the remaining functions.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
' -----------------------------------------------------------------------------------------------------------------------
Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
          Dim TwoDArray As Variant

1         On Error GoTo ErrHandler

2         Select Case NumDimensions(TheArray)
              Case 0
3                 ReDim TwoDArray(1 To 1, 1 To 1)
4                 TwoDArray(1, 1) = TheArray
5                 TheArray = TwoDArray
6                 NR = 1: NC = 1
7             Case 1
                  Dim i As Long
                  Dim LB As Long
8                 LB = LBound(TheArray, 1)
9                 NR = 1: NC = UBound(TheArray, 1) - LB + 1
10                ReDim TwoDArray(1 To 1, 1 To NC)
11                For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
12                    TwoDArray(1, i) = TheArray(LB + i - 1)
13                Next i
14                TheArray = TwoDArray
15            Case 2
16                NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
17                NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
                  'Nothing to do
18            Case Else
19                Throw "Cannot convert array of dimension greater than two"
20        End Select

21        Exit Sub
ErrHandler:
22        Throw "#Force2DArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
' -----------------------------------------------------------------------------------------------------------------------
Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
1         If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
2         Force2DArray RangeOrArray, NR, NC
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayRMulti
' Author    : Philip Swannell
' Date      : 11-May-2015
' Purpose   : Force many variables to 2-d array status
' -----------------------------------------------------------------------------------------------------------------------
Sub Force2DArrayRMulti(ParamArray RangesOrArrays())
          Dim i As Long
1         For i = LBound(RangesOrArrays, 1) To UBound(RangesOrArrays, 1)
2             Force2DArrayR RangesOrArrays(i)
3         Next i
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
' -----------------------------------------------------------------------------------------------------------------------
Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim y As Long
1         If Not IsArray(x) Then
2             NumDimensions = 0
3             Exit Function
4         Else
5             On Error GoTo ExitPoint
6             i = 1
7             Do While True
8                 y = LBound(x, i)
9                 i = i + 1
10            Loop
11        End If
ExitPoint:
12        NumDimensions = i - 1
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayRange
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Places arrays side by side. If the arrays are of unequal height then they will be padded
'             underneath with #NA! values.
' Arguments
' ArraysToRange:
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayRange(ParamArray ArraysToRange())
Attribute sArrayRange.VB_Description = "Places arrays side by side. If the arrays are of unequal height then they will be padded underneath with #N/A! values."
Attribute sArrayRange.VB_ProcData.VB_Invoke_Func = " \n24"

          Dim AllC As Long
          Dim AllR As Long
          Dim c As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim R As Long
          Dim ReturnArray()
          Dim Y0 As Long

1         On Error GoTo ErrHandler

          Static NA As Variant
2         If IsEmpty(NA) Then NA = CVErr(xlErrNA)

3         If IsMissing(ArraysToRange) Then
4             sArrayRange = CreateMissing()
5         Else
6             For i = LBound(ArraysToRange) To UBound(ArraysToRange)
7                 If TypeName(ArraysToRange(i)) = "Range" Then ArraysToRange(i) = ArraysToRange(i).Value
8                 If IsMissing(ArraysToRange(i)) Then
9                     R = 0: c = 0
10                Else
11                    Select Case NumDimensions(ArraysToRange(i))
                          Case 0
12                            R = 1: c = 1
13                        Case 1
14                            R = 1
15                            c = UBound(ArraysToRange(i)) - LBound(ArraysToRange(i)) + 1
16                        Case 2
17                            R = UBound(ArraysToRange(i), 1) - LBound(ArraysToRange(i), 1) + 1
18                            c = UBound(ArraysToRange(i), 2) - LBound(ArraysToRange(i), 2) + 1
19                    End Select
20                End If
21                If R > AllR Then AllR = R
22                AllC = AllC + c
23            Next i

24            If AllR = 0 Then
25                sArrayRange = CreateMissing
26                Exit Function
27            End If

28            ReDim ReturnArray(1 To AllR, 1 To AllC)

29            Y0 = 1
30            For i = LBound(ArraysToRange) To UBound(ArraysToRange)
31                If Not IsMissing(ArraysToRange(i)) Then
32                    Select Case NumDimensions(ArraysToRange(i))
                          Case 0
33                            R = 1: c = 1
34                            ReturnArray(1, Y0) = ArraysToRange(i)
35                        Case 1
36                            R = 1
37                            c = UBound(ArraysToRange(i)) - LBound(ArraysToRange(i)) + 1
38                            For j = 1 To c
39                                ReturnArray(1, Y0 + j - 1) = ArraysToRange(i)(j + LBound(ArraysToRange(i)) - 1)
40                            Next j
41                        Case 2
42                            R = UBound(ArraysToRange(i), 1) - LBound(ArraysToRange(i), 1) + 1
43                            c = UBound(ArraysToRange(i), 2) - LBound(ArraysToRange(i), 2) + 1

44                            For j = 1 To R
45                                For k = 1 To c
46                                    ReturnArray(j, Y0 + k - 1) = ArraysToRange(i)(j + LBound(ArraysToRange(i), 1) - 1, k + LBound(ArraysToRange(i), 2) - 1)
47                                Next k
48                            Next j

49                    End Select
50                    If R < AllR Then        'Pad with #NA! values
51                        For j = R + 1 To AllR
52                            For k = 1 To c
53                                ReturnArray(j, Y0 + k - 1) = NA
54                            Next k
55                        Next j
56                    End If

57                    Y0 = Y0 + c
58                End If
59            Next i
60            sArrayRange = ReturnArray
61        End If

62        Exit Function
ErrHandler:
63        sArrayRange = "#sArrayRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArraysIdentical
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Returns TRUE if the two input arrays are identical. That is, they are the same size and
'             shape and every pair of elements are equal.
'
' Arguments
' Array1    : The first array to compare.
' Array2    : The second array to compare.
' CaseSensitive: TRUE for case sensitive comparison of strings. FALSE or omitted for case insensitive
'             comparison.
' PermitBaseDifference: This argument is not relevant when using the function in an Excel formula and should be
'             omitted. If used from VBA code, then setting it to TRUE allows "zero-based"
'             arrays to be compared with "one-based" arrays.
' -----------------------------------------------------------------------------------------------------------------------
Function sArraysIdentical(ByVal Array1, ByVal Array2, Optional CaseSensitive As Boolean, Optional PermitBaseDifference As Boolean = False) As Variant
Attribute sArraysIdentical.VB_Description = "Returns TRUE if the two input arrays are identical. That is, they are the same size and shape and every pair of elements are equal.\n"
Attribute sArraysIdentical.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim cN As Long
          Dim i As Long
          Dim j As Long
          Dim rN As Long
1         On Error GoTo ErrHandler

2         Force2DArrayR Array1: Force2DArrayR Array2

3         If (UBound(Array1, 1) - LBound(Array1, 1)) <> (UBound(Array2, 1) - LBound(Array2, 1)) Then
4             sArraysIdentical = False
5         ElseIf (UBound(Array1, 2) - LBound(Array1, 2)) <> (UBound(Array2, 2) - LBound(Array2, 2)) Then
6             sArraysIdentical = False
7         Else
8             If Not PermitBaseDifference Then
9                 If (LBound(Array1, 1) <> LBound(Array2, 1)) Or (LBound(Array1, 2) <> LBound(Array2, 2)) Then
10                    sArraysIdentical = False
11                    Exit Function
12                End If
13            End If
14            rN = LBound(Array2, 1) - LBound(Array1, 1)
15            cN = LBound(Array2, 2) - LBound(Array1, 2)
16            For i = LBound(Array1, 1) To UBound(Array1, 1)
17                For j = LBound(Array1, 2) To UBound(Array1, 2)
18                    If Not sEquals(Array1(i, j), Array2(i + rN, j + cN), CaseSensitive) Then
19                        sArraysIdentical = False
20                        Exit Function
21                    End If
22                Next j
23            Next i
24            sArraysIdentical = True
25        End If

26        Exit Function
ErrHandler:
27        sArraysIdentical = "#sArraysIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArraysNearlyIdentical
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Returns TRUE if the two input arrays are nearly identical. That is, they are the same size
'             and shape and every pair of elements are equal, or in the case of pairs of
'             numbers, are nearly equal.
'
' Arguments
' Array1    : The first array to compare.
' Array2    : The second array to compare.
' CaseSensitive: TRUE for case sensitive comparison of strings. FALSE or omitted for case insensitive
'             comparison.
' Epsilon   : Epsilon determines the tolerance for comparison of two numbers via the formula:
'             A Nearly Equals B iff Abs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))
'             If omitted Epsilon defaults to 0.000000000000001 (i.e. 10^-15)
' PermitBaseDifference: This argument is not relevant when using the function in an Excel formula and should be
'             omitted. If used from VBA code, then setting it to TRUE allows "zero-based"
'             arrays to be compared with "one-based" arrays.
' -----------------------------------------------------------------------------------------------------------------------
Function sArraysNearlyIdentical(ByVal Array1, ByVal Array2, Optional CaseSensitive As Boolean, Optional Epsilon As Double = cEPSILON, Optional PermitBaseDifference As Boolean) As Variant
Attribute sArraysNearlyIdentical.VB_Description = "Returns TRUE if the two input arrays are nearly identical. That is, they are the same size and shape and every pair of elements are equal, or in the case of pairs of numbers, are nearly equal.\n"
Attribute sArraysNearlyIdentical.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim cN As Long
          Dim i As Long
          Dim j As Long
          Dim rN As Long
1         On Error GoTo ErrHandler

2         Force2DArrayR Array1: Force2DArrayR Array2

3         If (UBound(Array1, 1) - LBound(Array1, 1)) <> (UBound(Array2, 1) - LBound(Array2, 1)) Then
4             sArraysNearlyIdentical = False
5         ElseIf (UBound(Array1, 2) - LBound(Array1, 2)) <> (UBound(Array2, 2) - LBound(Array2, 2)) Then
6             sArraysNearlyIdentical = False
7         Else
8             If Not PermitBaseDifference Then
9                 If (LBound(Array1, 1) <> LBound(Array2, 1)) Or (LBound(Array1, 2) <> LBound(Array2, 2)) Then
10                    sArraysNearlyIdentical = False
11                    Exit Function
12                End If
13            End If

14            rN = LBound(Array2, 1) - LBound(Array1, 1)
15            cN = LBound(Array2, 2) - LBound(Array1, 2)
16            For i = LBound(Array1, 1) To UBound(Array1, 1)
17                For j = LBound(Array1, 2) To UBound(Array1, 2)
18                    If Not sNearlyEquals(Array1(i, j), Array2(i + rN, j + cN), CaseSensitive, Epsilon) Then
19                        sArraysNearlyIdentical = False
20                        Exit Function
21                    End If
22                Next j
23            Next i
24            sArraysNearlyIdentical = True
25        End If

26        Exit Function
ErrHandler:
27        sArraysNearlyIdentical = "#sArraysNearlyIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayStack
' Author    : Philip Swannell
' Date      : 25-Jun-2013
' Purpose   : Places arrays on top of one another. If the arrays are of unequal width then they will be
'             padded to the right with #NA! values.
' Arguments
' ArraysToStack:
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayStack(ParamArray ArraysToStack())
Attribute sArrayStack.VB_Description = "Places arrays on top of one another. If the arrays are of unequal width then they will be padded to the right with #N/A! values."
Attribute sArrayStack.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim AllC As Long
          Dim AllR As Long
          Dim c As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim R As Long
          Dim R0 As Long
          Dim ReturnArray()
1         On Error GoTo ErrHandler

          Static NA As Variant
2         If IsMissing(ArraysToStack) Then
3             sArrayStack = CreateMissing()
4         Else
5             If IsEmpty(NA) Then NA = CVErr(xlErrNA)

6             For i = LBound(ArraysToStack) To UBound(ArraysToStack)
7                 If TypeName(ArraysToStack(i)) = "Range" Then ArraysToStack(i) = ArraysToStack(i).Value
8                 If IsMissing(ArraysToStack(i)) Then
9                     R = 0: c = 0
10                Else
11                    Select Case NumDimensions(ArraysToStack(i))
                          Case 0
12                            R = 1: c = 1
13                        Case 1
14                            R = 1
15                            c = UBound(ArraysToStack(i)) - LBound(ArraysToStack(i)) + 1
16                        Case 2
17                            R = UBound(ArraysToStack(i), 1) - LBound(ArraysToStack(i), 1) + 1
18                            c = UBound(ArraysToStack(i), 2) - LBound(ArraysToStack(i), 2) + 1
19                    End Select
20                End If
21                If c > AllC Then AllC = c
22                AllR = AllR + R
23            Next i

24            If AllR = 0 Then
25                sArrayStack = CreateMissing
26                Exit Function
27            End If

28            ReDim ReturnArray(1 To AllR, 1 To AllC)

29            R0 = 1
30            For i = LBound(ArraysToStack) To UBound(ArraysToStack)
31                If Not IsMissing(ArraysToStack(i)) Then
32                    Select Case NumDimensions(ArraysToStack(i))
                          Case 0
33                            R = 1: c = 1
34                            ReturnArray(R0, 1) = ArraysToStack(i)
35                        Case 1
36                            R = 1
37                            c = UBound(ArraysToStack(i)) - LBound(ArraysToStack(i)) + 1
38                            For j = 1 To c
39                                ReturnArray(R0, j) = ArraysToStack(i)(j + LBound(ArraysToStack(i)) - 1)
40                            Next j
41                        Case 2
42                            R = UBound(ArraysToStack(i), 1) - LBound(ArraysToStack(i), 1) + 1
43                            c = UBound(ArraysToStack(i), 2) - LBound(ArraysToStack(i), 2) + 1

44                            For j = 1 To R
45                                For k = 1 To c
46                                    ReturnArray(R0 + j - 1, k) = ArraysToStack(i)(j + LBound(ArraysToStack(i), 1) - 1, k + LBound(ArraysToStack(i), 2) - 1)
47                                Next k
48                            Next j

49                    End Select
50                    If c < AllC Then
51                        For j = 1 To R
52                            For k = c + 1 To AllC
53                                ReturnArray(R0 + j - 1, k) = NA
54                            Next k
55                        Next j
56                    End If
57                    R0 = R0 + R
58                End If
59            Next i

60            sArrayStack = ReturnArray
61        End If
62        Exit Function
ErrHandler:
63        sArrayStack = "#sArrayStack (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnReverse
' Author    : Philip Swannell
' Date      : 20-Jun-2013
' Purpose   : Turns an array upside down. The function returns an array of the same size as the input
'             array with the vertical order reversed, so that the last row is now first,
'             and the first row is now last.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnReverse(TheArray As Variant)
Attribute sColumnReverse.VB_Description = "Turns an array upside down. The function returns an array of the same size as the input array with the vertical order reversed, so that the last row is now first, and the first row is now last."
Attribute sColumnReverse.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim Left As Long
          Dim NC As Long
          Dim NR As Long
          Dim ResultArray() As Variant
          Dim Top As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray, NR, NC

3         Top = LBound(TheArray, 1)
4         Left = LBound(TheArray, 2) - 1

5         ReDim ResultArray(1 To NR, 1 To NC)
6         For i = 1 To NR
7             For j = 1 To NC
8                 ResultArray(i, j) = TheArray(NR - i + Top, Left + j)
9             Next j
10        Next i

11        sColumnReverse = ResultArray

12        Exit Function
ErrHandler:
13        sColumnReverse = "#sColumnReverse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sConcatenateStrings
' Author    : Philip Swannell
' Date      : 18-May-2015
' Purpose   : The function takes a column of strings and concatenates the entries together, with a
'             Delimiter string between each pair, to give a single output string (or row
'             array of strings if TheStrings has multiple columns).
' Arguments
' TheStrings: An array of strings. If this array contains non-strings then those elements will be cast
'             to strings before concatenation is done.
' Delimiter : The delimiter character. If omitted defaults to a comma. Can be specified as multiple
'             characters or the empty string.
' -----------------------------------------------------------------------------------------------------------------------
Function sConcatenateStrings(ByVal TheStrings As Variant, Optional Delimiter As String = ",") As Variant
Attribute sConcatenateStrings.VB_Description = "The function takes a column of strings and concatenates the entries together, with a Delimiter string between each pair, to give a single output string (or row array of strings if TheStrings has multiple columns)."
Attribute sConcatenateStrings.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim ResultArray() As String
          Dim TempArray() As String
          Dim xO As Long
          Dim yO As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR TheStrings, NR, NC
3         xO = LBound(TheStrings, 1) - 1
4         yO = LBound(TheStrings, 2) - 1

5         ReDim ResultArray(1 To 1, 1 To NC)
6         ReDim TempArray(1 To NR)

7         For j = 1 To NC
8             For i = 1 To NR
9                 If VarType(TheStrings(xO + i, yO + j)) <> vbString Then
10                    TheStrings(xO + i, yO + j) = NonStringToString(TheStrings(xO + i, yO + j))
11                End If
12                TempArray(i) = TheStrings(xO + i, yO + j)
13            Next i
14            ResultArray(1, j) = VBA.Join(TempArray, Delimiter)
15        Next j

16        If NC = 1 Then
17            sConcatenateStrings = ResultArray(1, 1)
18        Else
19            sConcatenateStrings = ResultArray
20        End If

21        Exit Function
ErrHandler:
22        sConcatenateStrings = "#sConcatenateStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCountDistinctItems
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Returns a two-column report on the number of occurrences of each distinct item within
'             TheArray. First column gives the item itself, the second the number of times
'             that item appears. The return is sorted in descending order on the number of
'             appearances.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sCountDistinctItems(ByVal TheArray As Variant)
Attribute sCountDistinctItems.VB_Description = "Returns a two-column report on the number of occurrences of each distinct item within TheArray. First column gives the item itself, the second the number of times that item appears. The return is sorted in descending order on the number of appearances."
Attribute sCountDistinctItems.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim NC As Long
          Dim NR As Long
          Dim TmpArray As Variant
          Dim UseExcelSortMethod As Boolean

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray, NR, NC
3         If NC <> 1 Then
4             TheArray = sReshape(TheArray, sNRows(TheArray) * sNCols(TheArray), 1)
5         End If
6         UseExcelSortMethod = ExcelSupportsSpill()

7         If Not UseExcelSortMethod Then
8             UseExcelSortMethod = TypeName(Application.Caller) = "Error"
9         End If

10        TmpArray = sSortedArray(TheArray, , , , , , , , UseExcelSortMethod)
11        TmpArray = sCountRepeats(TmpArray, "CH")
12        TmpArray = sSortedArray(TmpArray, 2, , , False, , , , UseExcelSortMethod)
13        sCountDistinctItems = TmpArray

14        Exit Function
ErrHandler:
15        sCountDistinctItems = "#sCountDistinctItems (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCountRepeats
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Reports on consecutive repeats in a column of data. Template can contain characters C, F,
'             T or H. C for Column, F for From, T for To and H for How Many.
'             Example:
'             Column =
'                X
'                X
'                Y
'                Y
'                Y
'                X
'             Template = "CH"
'             Return =
'             X   2
'             Y   3
'             X   1
' Arguments
' Column    : A single column array of arbitrary data - numbers, strings, logical values, etc.
' Template  : Characters of Template determine the columns returned
'             C - the element
'             F - 1st row number of a consecutive sequence of repeats
'             T - last row row number of a consecutive sequence of repeats
'             H - How many consecutive repeats - 1 if an element is not repeated
' -----------------------------------------------------------------------------------------------------------------------
Function sCountRepeats(Column, Template As String)
Attribute sCountRepeats.VB_Description = "Reports on consecutive repeats in a column of data. Template can contain characters C, F, T or H. C for Column, F for From, T for To and H for How Many.\nExample:\nColumn = \n    X\n    X\n    Y\n    Y\n    Y\n    X\nTemplate = ""CH""\nReturn =\nX   2\nY   3\nX   1"
Attribute sCountRepeats.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim c As Variant
          Dim F As Long
          Dim H As Long
          Dim i As Long
          Dim j As Long
          Dim objC As clsStacker
          Dim objF As clsStacker
          Dim objH As clsStacker
          Dim objT As clsStacker
          Dim sNRows As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR Column

          'Check Inputs
3         If sNCols(Column) > 1 Then
4             sCountRepeats = "#Column must be a single column array!"
5             Exit Function
6         End If
7         For i = 1 To Len(Template)
8             Select Case UCase$(Mid$(Template, i, 1))
                  Case "C", "F", "T", "H"
                      'OK
9                 Case Else
10                    sCountRepeats = "#Unrecognised character in Template. Allowed characters: " & _
                          """C"" (for Column), ""H"" (for How many), ""F"" (for From) or ""T"" (for To)!"
11                    Exit Function
12            End Select
13        Next i

14        Set objC = CreateStacker()
15        Set objF = CreateStacker()
16        Set objT = CreateStacker()
17        Set objH = CreateStacker()

18        c = Column(LBound(Column, 1), 1)
19        F = LBound(Column, 1)
20        H = 1
21        objC.Stack0D c
22        objF.Stack0D F
23        sNRows = sNRows + 1
24        For i = LBound(Column, 1) + 1 To UBound(Column, 1)
25            If sEquals(Column(i - 1, 1), Column(i, 1)) Then
26                H = H + 1
27            Else
28                c = Column(i, 1)
29                F = i
30                objC.Stack0D c
31                sNRows = sNRows + 1
32                objF.Stack0D F
33                objH.Stack0D H
34                objT.Stack0D i - 1
35                H = 1
36            End If
37        Next i
38        objH.Stack0D H
39        objT.Stack0D UBound(Column, 1)

          'Construct return
          Dim TmpArray() As Variant
          Dim TmpCol As Variant
40        ReDim TmpArray(1 To sNRows, 1 To Len(Template))
41        For i = 1 To Len(Template)
42            Select Case UCase$(Mid$(Template, i, 1))
                  Case "C"
43                    TmpCol = objC.ReportInTranspose
44                Case "F"
45                    TmpCol = objF.ReportInTranspose
46                Case "H"
47                    TmpCol = objH.ReportInTranspose
48                Case "T"
49                    TmpCol = objT.ReportInTranspose
50            End Select
51            For j = 1 To sNRows
52                TmpArray(j, i) = TmpCol(1, j)
53            Next j
54        Next i

55        sCountRepeats = TmpArray

56        Exit Function
ErrHandler:
57        sCountRepeats = "#sCountRepeats (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDrop
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Removes the first (or last) few rows of an array. The remaining rows are returned. If
'             NumToDrop is greater than or equal to the height of TheArray an error is
'             returned.
' Arguments
' TheArray  : An array of arbitrary values.
' NumToDrop : Integer number of rows to remove. If positive, top rows are removed, and if negative,
'             bottom rows are removed.
' -----------------------------------------------------------------------------------------------------------------------
Function sDrop(TheArray, NumToDrop As Long)
Attribute sDrop.VB_Description = "Removes the first (or last) few rows of an array. The remaining rows are returned. If NumToDrop is greater than or equal to the height of TheArray an error is returned."
Attribute sDrop.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim NR As Long
1         NR = sNRows(TheArray)
2         If NumToDrop = 0 Then
3             sDrop = TheArray
4         ElseIf NumToDrop >= NR Or NumToDrop <= -NR Then
5             sDrop = "#Cannot drop that many!"
6         ElseIf NumToDrop > 0 Then
7             sDrop = sSubArray(TheArray, NumToDrop + 1)
8         Else
9             sDrop = sSubArray(TheArray, , , NR + NumToDrop)
10        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFirstError
' Author    : Philip Swannell
' Date      : 03-Apr-2019
' Purpose   : Returns the first error string found in an array, or the string "No errors found". An
'             error string is a string starting with '#' and ending with '!'.
' Arguments
' x         : A single value or an array of values.
' -----------------------------------------------------------------------------------------------------------------------
Function sFirstError(x As Variant) As String
Attribute sFirstError.VB_Description = "Returns the first error string found in an array, or the string ""No errors found"". An error string is a string starting with '#' and ending with '!'."
Attribute sFirstError.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim tmp As Variant
1         On Error GoTo ErrHandler
2         If IsArray(x) Then
3             For Each tmp In x
4                 If sIsErrorString(tmp) Then
5                     sFirstError = CStr(tmp)
6                     Exit Function
7                 End If
8             Next tmp
9             sFirstError = "No errors found"
10        ElseIf sIsErrorString(x) Then
11            sFirstError = CStr(x)
12        Else
13            sFirstError = "No errors found"
14        End If

15        Exit Function
ErrHandler:
16        sFirstError = "#sFirstError (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIsErrorString
' Author    : Philip Swannell
' Date      : 02-May-2016
' Purpose   : Returns TRUE if x is a string starting with "#" and ending with "!". If X is an array then
'             the return will be FALSE unless x has only one element and that element is a
'             string that starts with "#" and ends with "!"
' Arguments
' X         : The value to test.
' -----------------------------------------------------------------------------------------------------------------------
Function sIsErrorString(ByVal x As Variant) As Boolean
Attribute sIsErrorString.VB_Description = "Returns TRUE if x is a string starting with ""#"" and ending with ""!"". If X is an array then the return will be FALSE unless x has only one element and that element is a string that starts with ""#"" and ends with ""!"""
Attribute sIsErrorString.VB_ProcData.VB_Invoke_Func = " \n25"
1         On Error GoTo ErrHandler
2         If TypeName(x) = "Range" Then x = x.Value2
3         If VarType(x) = vbString Then
4             If Left$(x, 1) = "#" Then
5                 If Right$(x, 1) = "!" Then
6                     sIsErrorString = True
7                 End If
8             End If
9         ElseIf IsArray(x) Then
10            Select Case NumDimensions(x)
                  Case 1
11                    If UBound(x) = LBound(x) Then
12                        sIsErrorString = sIsErrorString(x(LBound(x)))
13                    End If
14                Case 2
15                    If UBound(x, 1) = LBound(x, 1) Then
16                        If UBound(x, 2) = LBound(x, 2) Then
17                            sIsErrorString = sIsErrorString(x(LBound(x, 1), LBound(x, 2)))
18                        End If
19                    End If
20            End Select
21        End If
22        Exit Function
ErrHandler:
23        Throw "#sIsErrorString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
