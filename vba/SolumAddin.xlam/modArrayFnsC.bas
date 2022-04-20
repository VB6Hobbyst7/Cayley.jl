Attribute VB_Name = "modArrayFnsC"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumber
' Author    : Philip Swannell
' Date      : 05-May-2015
' Purpose   : Is a singleton a number?
' -----------------------------------------------------------------------------------------------------------------------
Function IsNumber(x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong, 20 'vbLongLong
2                 IsNumber = True
3         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumberOrDate
' Author    : Philip Swannell
' Date      : 05-May-2015
' Purpose   : Is a singleton a number or date
' -----------------------------------------------------------------------------------------------------------------------
Function IsNumberOrDate(x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong, vbDate, 20 'vbLongLong
2                 IsNumberOrDate = True
3         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsWholeNumber
' Author    : Philip Swannell
' Date      : 21-Apr-2017
' Purpose   : is a singleton a whole number?
' -----------------------------------------------------------------------------------------------------------------------
Function IsWholeNumber(x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbInteger, vbSingle, vbLong, 20 ' vbLongLong
2                 IsWholeNumber = True
3             Case vbDouble
4                 IsWholeNumber = (CLng(x) = x)
5         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayAbs
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Element-wise absolute value of an array of numbers.
' Arguments
' TheArray  : Any number or array of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayAbs(ByVal TheArray)
Attribute sArrayAbs.VB_Description = "Element-wise absolute value of an array of numbers."
Attribute sArrayAbs.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = SafeAbs(TheArray(i, j))
7             Next j
8         Next i
9         sArrayAbs = Result

10        Exit Function
ErrHandler:
11        sArrayAbs = "#sArrayAbs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayCount
' Author    : Philip Swannell
' Date      : 30-Sep-2013
' Purpose   : Returns the integer count of the number of elements of TheArray which are TRUE. All other
'             values, including error values, are ignored.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayCount(TheArray As Variant)
Attribute sArrayCount.VB_Description = "Returns the integer count of the number of elements of TheArray which are TRUE. All other values, including error values, are ignored."
Attribute sArrayCount.VB_ProcData.VB_Invoke_Func = " \n24"
1         On Error GoTo ErrHandler
          Dim Res As Long
          Dim x As Variant

2         Force2DArrayR TheArray

3         For Each x In TheArray
4             If VarType(x) = vbBoolean Then If x Then Res = Res + 1
5         Next

6         sArrayCount = Res

7         Exit Function
ErrHandler:
8         sArrayCount = "#sArrayCount (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayExp
' Author    : Philip Swannell
' Date      : 27-Apr-2015
' Purpose   : Returns the exponential function (natural to base e) of the input numeric array, exp(x).
'             Non-numbers in the input  will produce an error string (e.g. "#Type
'             mismatch!") in the corresponding element of the output array.
' Arguments
' TheArray  : An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayExp(ByVal TheArray)
Attribute sArrayExp.VB_Description = "Returns the exponential function (natural to base e) of the input numeric array, exp(x). Non-numbers in the input  will produce an error string (e.g. ""#Type mismatch!"") in the corresponding element of the output array."
Attribute sArrayExp.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = SafeExp(TheArray(i, j))
7             Next j
8         Next i
9         sArrayExp = Result
10        Exit Function
ErrHandler:
11        sArrayExp = "#sArrayExp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayFind
' Author    : Philip Swannell
' Date      : 15-Nov-2013
' Purpose   : Returns TRUE if searchFor is a sub-string of SearchWithin, with case-insensitive matching.
'             SearchWithin may be an array of strings
' Arguments
' searchFor : The string to search for - case insensitive.
' SearchWithin: The string to search within. May be an array of strings in which case the return from the
'             function is also an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayFind(SearchFor As String, ByVal SearchWithin As Variant, Optional Mode As Long = 0)
Attribute sArrayFind.VB_Description = "Tests if searchFor is a sub-string of SearchWithin, with case-insensitive matching. SearchWithin may be an array of strings"
Attribute sArrayFind.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res() As Variant
1         On Error GoTo ErrHandler

2         Force2DArrayR SearchWithin, NR, NC

3         ReDim Res(1 To NR, 1 To NC)
4         Select Case Mode
              Case 0
                  'Return True if found, false o.w.
5                 For j = 1 To NC
6                     For i = 1 To NR
7                         Res(i, j) = InStr(1, CStr(SearchWithin(i, j)), SearchFor, vbTextCompare) > 0
8                     Next i
9                 Next j
10            Case 1
                  'Return position of first match, reading from the left, or 0 if not found
11                For j = 1 To NC
12                    For i = 1 To NR
13                        Res(i, j) = InStr(1, CStr(SearchWithin(i, j)), SearchFor, vbTextCompare)
14                    Next i
15                Next j
16            Case -1
                  'Return position of first match, reading from the right, or 0 if not found
17                For j = 1 To NC
18                    For i = 1 To NR
19                        Res(i, j) = InStrRev(CStr(SearchWithin(i, j)), SearchFor, , vbTextCompare)
20                    Next i
21                Next j
22            Case Else
23                Throw "Mode must be 0, 1 or -1"
24        End Select

25        sArrayFind = Res

26        Exit Function
ErrHandler:
27        Throw "#sArrayFind (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayIsLogical
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Performs an elementwise test on TheArray to see if its entries are logical.
' Arguments
' TheArray  : Array of elements to be tested.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayIsLogical(ByVal TheArray)
Attribute sArrayIsLogical.VB_Description = "Performs an element-wise test on TheArray to see if its entries are logical."
Attribute sArrayIsLogical.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = VarType(TheArray(i, j)) = vbBoolean
7             Next j
8         Next i
9         sArrayIsLogical = Result
10        Exit Function
ErrHandler:
11        sArrayIsLogical = "#sArrayIsLogical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayIsNonTrivialText
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Performs an elementwise test on TheArray to see if its entries are text with at least one
'             character.
' Arguments
' TheArray  : Array of elements to be tested.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayIsNonTrivialText(ByVal TheArray)
Attribute sArrayIsNonTrivialText.VB_Description = "Performs an element-wise test on TheArray to see if its entries are text with at least one character."
Attribute sArrayIsNonTrivialText.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(False, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 If VarType(TheArray(i, j)) = vbString Then
7                     Result(i, j) = (Len(TheArray(i, j)) > 0)
8                 End If
9             Next j
10        Next i
11        sArrayIsNonTrivialText = Result
12        Exit Function
ErrHandler:
13        sArrayIsNonTrivialText = "#sArrayIsNonTrivialText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayIsNumber
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Performs an elementwise test on TheArray to see if its entries are numeric.
' Arguments
' TheArray  : Array of elements to be tested.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayIsNumber(ByVal TheArray)
Attribute sArrayIsNumber.VB_Description = "Performs an element-wise test on TheArray to see if its entries are numeric."
Attribute sArrayIsNumber.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = IsNumber(TheArray(i, j))
7             Next j
8         Next i
9         sArrayIsNumber = Result
10        Exit Function
ErrHandler:
11        sArrayIsNumber = "#sArrayIsNumber (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayIsText
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Performs an elementwise test on TheArray to see if its entries are text.
' Arguments
' TheArray  : Array of elements to be tested.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayIsText(ByVal TheArray)
Attribute sArrayIsText.VB_Description = "Performs an element-wise test on TheArray to see if its entries are text."
Attribute sArrayIsText.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = VarType(TheArray(i, j)) = vbString
7             Next j
8         Next i
9         sArrayIsText = Result
10        Exit Function
ErrHandler:
11        sArrayIsText = "#sArrayIsText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayLog
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Element-wise natural log of an array of numbers.
' Arguments
' TheArray  : An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayLog(ByVal TheArray)
Attribute sArrayLog.VB_Description = "Element-wise natural log of an array of numbers."
Attribute sArrayLog.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 Result(i, j) = SafeLog(TheArray(i, j))
7             Next j
8         Next i
9         sArrayLog = Result

10        Exit Function
ErrHandler:
11        sArrayLog = "#sArrayLog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayNot
' Author    : Philip Swannell
' Date      : 09-May-2015
' Purpose   : Performs element-wise negation of an array of values. For logical values FALSE becomes
'             TRUE and vice-versa. For non-logical values an error string is returned.
' Arguments
' TheArray  : The array to negate.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayNot(ByVal TheArray)
Attribute sArrayNot.VB_Description = "Element-wise negation of an array of logical values. Non-logical values yield an error string embedded in the return."
Attribute sArrayNot.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Result = sReshape(0, N, M)
4         For i = 1 To N
5             For j = 1 To M
6                 If VarType(TheArray(i, j)) = vbBoolean Then
7                     Result(i, j) = Not (TheArray(i, j))
8                 Else
9                     Result(i, j) = "#Non Boolean detected!"
10                End If
11            Next j
12        Next i
13        sArrayNot = Result
14        Exit Function
ErrHandler:
15        sArrayNot = "#sArrayNot (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArraySquare
' Author    : Philip Swannell
' Date      : 24-May-2015
' Purpose   : Puts together four arrays into a square. If the input arrays size do not fit, arrays will
'             be padded with #NA to adjust.
'
'             The function is equivalent to:
'             sArrayStack(sArrayRange(TL, TR), sArrayRange(BL, BR)) though roughly 15%
'             faster.
' Arguments
' TopLeft   : Array for top-left quadrant
' TopRight  : Array for top-right quadrant
' BottomLeft: Array for bottom-left quadrant
' BottomRight: Array for bottom-right quadrant
' -----------------------------------------------------------------------------------------------------------------------
Function sArraySquare(Optional TopLeft, Optional TopRight, Optional BottomLeft, Optional BottomRight)
Attribute sArraySquare.VB_Description = "Puts together four arrays into a square. If the input arrays size do not fit, arrays will be padded with #NA to adjust.\n\nThe function is equivalent to:\nsArrayStack(sArrayRange(TL, TR), sArrayRange(BL, BR)) though roughly 15% faster."
Attribute sArraySquare.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim BLnC As Long
          Dim BLnR As Long
          Dim BLoC As Long
          Dim BLoR As Long
          Dim BRnC As Long
          Dim BRnR As Long
          Dim BRoC As Long
          Dim BRoR As Long
          Dim i As Long
          Dim j As Long
          Dim RC As Long
          Dim Result As Variant
          Dim ROffset As Long
          Dim RR As Long
          Dim SomeArgs As Boolean
          Dim TLnC As Long
          Dim TLnR As Long
          Dim TLoC As Long
          Dim TLoR As Long
          Dim TRnC As Long
          Dim TRnR As Long
          Dim TRoC As Long
          Dim TRoR As Long

1         On Error GoTo ErrHandler

2         If Not IsMissing(TopLeft) Then
3             Force2DArrayR TopLeft, TLnR, TLnC
4             TLoR = LBound(TopLeft, 1) - 1
5             TLoC = LBound(TopLeft, 2) - 1
6             SomeArgs = True
7         End If
8         If Not IsMissing(TopRight) Then
9             Force2DArrayR TopRight, TRnR, TRnC
10            TRoR = LBound(TopRight, 1) - 1
11            TRoC = LBound(TopRight, 2) - 1
12            SomeArgs = True
13        End If
14        If Not IsMissing(BottomLeft) Then
15            Force2DArrayR BottomLeft, BLnR, BLnC
16            BLoR = LBound(BottomLeft, 1) - 1
17            BLoC = LBound(BottomLeft, 2) - 1
18            SomeArgs = True
19        End If
20        If Not IsMissing(BottomRight) Then
21            Force2DArrayR BottomRight, BRnR, BRnC
22            BRoR = LBound(BottomRight, 1) - 1
23            BRoC = LBound(BottomRight, 2) - 1
24            SomeArgs = True
25        End If

26        If Not SomeArgs Then Throw "No arguments supplied"

27        ROffset = SafeMax(TLnR, TRnR)
28        RR = ROffset + SafeMax(BLnR, BRnR)
29        RC = SafeMax(TLnC + TRnC, BLnC + BRnC)

30        Result = sReshape(CVErr(xlErrNA), RR, RC)

31        For i = 1 To TLnR
32            For j = 1 To TLnC
33                Result(i, j) = TopLeft(TLoR + i, TLoC + j)
34            Next j
35        Next i
36        For i = 1 To TRnR
37            For j = 1 To TRnC
38                Result(i, j + TLnC) = TopRight(TRoR + i, TRoC + j)
39            Next j
40        Next i
41        For i = 1 To BLnR
42            For j = 1 To BLnC
43                Result(ROffset + i, j) = BottomLeft(BLoR + i, BLoC + j)
44            Next j
45        Next i
46        For i = 1 To BRnR
47            For j = 1 To BRnC
48                Result(ROffset + i, j + BLnC) = BottomRight(BRoR + i, BRoC + j)
49            Next j
50        Next i

51        sArraySquare = Result

52        Exit Function
ErrHandler:
53        sArraySquare = "#sArraySquare (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub sArraySquareSpeedTest()
          Dim BottomLeft As Variant
          Dim BottomRight
          Dim NC As Long
          Dim NR As Long
          Dim Res1
          Dim Res2
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim TopLeft As Variant
          Dim TopRight As Variant
1         NR = 1000
2         NC = 100

3         TopLeft = sReshape(1, NR, NC)
4         TopRight = sReshape("Foo", NR, NC)
5         BottomLeft = sReshape(0, NR, NC)
6         BottomRight = sReshape(True, NR, NC)

7         t1 = sElapsedTime()
8         Res1 = sArraySquare(TopLeft, TopRight, BottomLeft, BottomRight)
9         t2 = sElapsedTime()
10        Res2 = sArrayStack(sArrayRange(TopLeft, TopRight), sArrayRange(BottomLeft, BottomRight))
11        t3 = sElapsedTime()
12        Debug.Print sArraysIdentical(Res1, Res2), t2 - t1, t3 - t2, (t2 - t1) / (t3 - t2)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayTranspose
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Returns the transpose of an array. 1 dimensional input becomes 2-d with 1 column. Output is always 1-based.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayTranspose(ByVal TheArray As Variant)
Attribute sArrayTranspose.VB_Description = "Returns the transpose of an array."
Attribute sArrayTranspose.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim Co As Long
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
          Dim Ro As Long
1         On Error GoTo ErrHandler
2         If TypeName(TheArray) = "Range" Then TheArray = TheArray.Value2
3         Select Case NumDimensions(TheArray)
              Case 0
4                 ReDim Result(1 To 1, 1 To 1)
5                 Result(1, 1) = TheArray
6             Case 1
7                 M = UBound(TheArray) - LBound(TheArray) + 1
8                 ReDim Result(1 To M, 1 To 1)
9                 Ro = LBound(TheArray) - 1
10                For i = 1 To M
11                    Result(i, 1) = TheArray(i + Ro)
12                Next i
13            Case 2
14                N = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
15                M = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
16                Ro = LBound(TheArray, 1) - 1
17                Co = LBound(TheArray, 2) - 1
18                ReDim Result(1 To M, 1 To N)
19                For i = 1 To N
20                    For j = 1 To M
21                        Result(j, i) = TheArray(i + Ro, j + Co)
22                    Next j
23                Next i
24            Case Else
25                Throw "Cannot transpose array with more than 2 dimensions"
26        End Select
27        sArrayTranspose = Result
28        Exit Function
ErrHandler:
29        sArrayTranspose = "#sArrayTranspose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMaxByChunks
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : For a column array of numbers, returns a column of numbers with the ith element of the
'             output equal to the maximum of the ith "chunk" of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' ChunkSize : A positive integer to define the size of the "chunks". E.g. if ChunklSize is 2 then the
'             ith output element will be maximum of the ith pair of elements of the input
'             column.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMaxByChunks(ArrayOfNumbers, ChunkSize As Long)
Attribute sColumnMaxByChunks.VB_Description = "For a column array of numbers, returns a column of numbers with the ith element of the output equal to the maximum of the ith ""chunk"" of the input. "
Attribute sColumnMaxByChunks.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         sColumnMaxByChunks = sColumnProcessByChunks(ArrayOfNumbers, ChunkSize, FuncIdMax)
3         Exit Function
ErrHandler:
4         sColumnMaxByChunks = "#sColumnMaxByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMeanByChunks
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : For a column array of numbers, returns a column of numbers with the ith element of the
'             output equal to the mean of the ith "chunk" of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' ChunkSize : A positive integer to define the size of the "chunks". E.g. if ChunklSize is 2 then the
'             ith output element will be average of the ith pair of elements of the input
'             column.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMeanByChunks(ArrayOfNumbers, ChunkSize As Long)
Attribute sColumnMeanByChunks.VB_Description = "For a column array of numbers, returns a column of numbers with the ith element of the output equal to the mean of the ith ""chunk"" of the input. "
Attribute sColumnMeanByChunks.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         sColumnMeanByChunks = sColumnProcessByChunks(ArrayOfNumbers, ChunkSize, FuncIdMean)
3         Exit Function
ErrHandler:
4         sColumnMeanByChunks = "#sColumnMeanByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMedian
' Author    : Philip Swannell
' Date      : 15-Jul-2017
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             return is the median of the corresponding column of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' IgnoreErrorValues: If FALSE (the default) then error values (#N/A!, #REF! etc) in ArrayOfNumbers yield error
'             strings in the corresponding element of the return. If TRUE error values are
'             excluded from calculation of the column medians. String values are always
'             excluded.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMedian(ByVal ArrayOfNumbers, Optional IgnoreErrors As Boolean)
Attribute sColumnMedian.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the return is the median of the corresponding column of the input."
Attribute sColumnMedian.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As Variant
2         Force2DArrayR ArrayOfNumbers, NR, NC
          Dim tmp() As Variant

3         ReDim Result(1 To 1, 1 To NC)
4         ReDim tmp(1 To NR)
5         For j = 1 To NC
6             If IgnoreErrors Then
7                 For i = 1 To NR
8                     If IsNumberOrDate(ArrayOfNumbers(i, j)) Then
9                         tmp(i) = ArrayOfNumbers(i, j)
10                    Else
11                        tmp(i) = vbNullString
12                    End If
13                Next i
14            Else
15                For i = 1 To NR
16                    tmp(i) = ArrayOfNumbers(i, j)
17                Next i
18            End If
19            Result(1, j) = SafeMedian(tmp)
20        Next j
21        sColumnMedian = Result
22        Exit Function
ErrHandler:
23        sColumnMedian = "#sColumnMedian (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMedianByChunks
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : For a column array of numbers, returns a column of numbers with the ith element of the
'             output equal to the median of the ith "chunk" of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' ChunkSize : A positive integer to define the size of the "chunks". E.g. if ChunklSize is 3 then the
'             ith output element will be median of the ith triplet of elements of the input
'             column.
' IgnoreNonNumbers: If FALSE (the default), non-numbers in the input handled as per Excel's MEDIAN: strings
'             and logical values are ignored, errors yield errors in the output. If TRUE,
'             then all non-numbers, including errors, in the input are ignored for
'             calculating medians.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMedianByChunks(ArrayOfNumbers, ChunkSize As Long, IgnoreNonNumbers As Boolean)
Attribute sColumnMedianByChunks.VB_Description = "For a column array of numbers, returns a column of numbers with the ith element of the output equal to the median of the ith ""chunk"" of the input. "
Attribute sColumnMedianByChunks.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim Chunk
          Dim Height As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim M As Long
          Dim Nin As Long
          Dim nOut As Long
          Dim Result() As Variant
          Dim StartRow As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR ArrayOfNumbers
3         M = sNCols(ArrayOfNumbers)
4         Nin = sNRows(ArrayOfNumbers)
5         nOut = ((Nin - 1) \ ChunkSize) + 1
6         ReDim Result(1 To nOut, 1 To M)

7         For j = 1 To M
8             For i = 1 To nOut
9                 StartRow = 1 + (i - 1) * ChunkSize
10                If i = nOut Then
11                    Height = Nin - StartRow + 1
12                Else
13                    Height = ChunkSize
14                End If
15                Chunk = sSubArray(ArrayOfNumbers, StartRow, j, Height, 1)
16                If IgnoreNonNumbers Then
17                    For k = 1 To Height
18                        If Not (IsNumber(Chunk(k, 1))) Then
19                            Chunk(k, 1) = Empty
20                        End If
21                    Next k
22                End If
23                Result(i, j) = SafeMedian(Chunk)
24            Next i
25        Next j

26        sColumnMedianByChunks = Result

27        Exit Function
ErrHandler:
28        sColumnMedianByChunks = "#sColumnMedianByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMinByChunks
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : For a column array of numbers, returns a column of numbers with the ith element of the
'             output equal to the minimum of the ith "chunk" of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' ChunkSize : A positive integer to define the size of the "chunks". E.g. if ChunklSize is 2 then the
'             ith output element will be minimum of the ith pair of elements of the input
'             column.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMinByChunks(ArrayOfNumbers, ChunkSize As Long)
Attribute sColumnMinByChunks.VB_Description = "For a column array of numbers, returns a column of numbers with the ith element of the output equal to the minimum of the ith ""chunk"" of the input. "
Attribute sColumnMinByChunks.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         sColumnMinByChunks = sColumnProcessByChunks(ArrayOfNumbers, ChunkSize, FuncIdMin)
3         Exit Function
ErrHandler:
4         sColumnMinByChunks = "#sColumnMinByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnProcessByChunks
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : "Core" function for sColumnMaxByChunks, sColumnMinByChunks, sColumnSumByChunks
' -----------------------------------------------------------------------------------------------------------------------
Private Function sColumnProcessByChunks(ArrayOfNumbers, ChunkSize As Long, Operator As BroadcastFuncID)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim M As Long
          Dim Nin As Long
          Dim nOut As Long
          Dim Result() As Variant
          Dim Temp As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR ArrayOfNumbers, Nin, M

3         If ChunkSize <= 0 Then
4             Throw "ChunkSize must be positive"
5         ElseIf ChunkSize = 1 Then
6             sColumnProcessByChunks = ArrayOfNumbers
7             Exit Function
8         End If

9         nOut = ((Nin - 1) \ ChunkSize) + 1

10        ReDim Result(1 To nOut, 1 To M)

11        For j = 1 To M
12            k = 1
13            Temp = ArrayOfNumbers(1, j)
14            For i = 2 To Nin
15                If i Mod ChunkSize = 1 Then
16                    If Operator = FuncIdMean Then
17                        Result(k, j) = SafeDivide(Temp, ChunkSize)
18                    Else
19                        Result(k, j) = Temp
20                    End If
21                    Temp = ArrayOfNumbers(i, j)
22                    k = k + 1
23                Else
24                    If Operator = FuncIdMax Then
25                        Temp = SafeMax(Temp, ArrayOfNumbers(i, j))
26                    ElseIf Operator = FuncIdMin Then
27                        Temp = SafeMin(Temp, ArrayOfNumbers(i, j))
28                    ElseIf Operator = FuncIdAdd Then
29                        Temp = SafeAdd(Temp, ArrayOfNumbers(i, j))
30                    ElseIf Operator = FuncIdMean Then
31                        Temp = SafeAdd(Temp, ArrayOfNumbers(i, j))
32                    Else
33                        Throw "Unexpected error: Operator not recognised"
34                    End If
35                End If
36            Next i
37            If Operator = FuncIdMean Then
38                Result(nOut, j) = SafeDivide(Temp, (((Nin - 1) Mod ChunkSize) + 1))
39            Else
40                Result(nOut, j) = Temp
41            End If
42        Next j

43        sColumnProcessByChunks = Result

44        Exit Function
ErrHandler:
45        Throw "#sColumnProcessByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnSumByChunks
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : For a column array of numbers, returns a column of numbers with the ith element of the
'             output equal to the sum of the ith "chunk" of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' ChunkSize : A positive integer to define the size of the "chunks". E.g. if ChunklSize is 2 then the
'             ith output element will be sum of the ith pair of elements of the input
'             column.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnSumByChunks(ArrayOfNumbers, ChunkSize As Long)
Attribute sColumnSumByChunks.VB_Description = "For a column array of numbers, returns a column of numbers with the ith element of the output equal to the sum of the ith ""chunk"" of the input. "
Attribute sColumnSumByChunks.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         sColumnSumByChunks = sColumnProcessByChunks(ArrayOfNumbers, ChunkSize, FuncIdAdd)
3         Exit Function
ErrHandler:
4         sColumnSumByChunks = "#sColumnSumByChunks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
