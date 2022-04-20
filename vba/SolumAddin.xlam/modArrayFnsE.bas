Attribute VB_Name = "modArrayFnsE"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIndex
' Author    : Philip Swannell
' Purpose   : Returns particular elements of TheArray, indexed
'             by RowNumbersOrHeaders and ColumnNumbersOrHeaders.
' Arguments
' TheArray  : Array of arbitrary values
' RowNumbersOrHeaders: Either an integer row index between 1 and the number of rows in TheArray or else a string
'             that matches a string appearing in the leftmost column of TheArray. This
'             argument may be an array.
' ColumnNumbersOrHeaders: Either an integer column index between 1 and the number of columns in TheArray or else a
'             string that matches a string appearing in the top row of TheArray. This
'             argument may be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sIndex(ByVal TheArray As Variant, Optional ByVal RowNumbersOrHeaders As Variant, Optional ByVal ColumnNumbersOrHeaders As Variant)
Attribute sIndex.VB_Description = "Returns particular elements of TheArray, indexed by RowNumbersOrHeaders and ColumnNumbersOrHeaders."
Attribute sIndex.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim c As Long
          Dim CHasStrings As Boolean
          Dim ColLock(1 To 2) As Boolean
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim R As Long
          Dim Result()
          Dim RHasStrings As Boolean
          Dim RowLock(1 To 2) As Boolean
          Dim x As Variant

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray
3         If IsMissing(RowNumbersOrHeaders) Then RowNumbersOrHeaders = sIntegers(sNRows(TheArray))
4         If IsMissing(ColumnNumbersOrHeaders) Then ColumnNumbersOrHeaders = sArrayTranspose(sIntegers(sNCols(TheArray)))

5         Force2DArrayRMulti RowNumbersOrHeaders, ColumnNumbersOrHeaders

6         For Each x In RowNumbersOrHeaders
7             If VarType(x) = vbString Then
8                 RHasStrings = True
9                 Exit For
10            End If
11        Next
12        If RHasStrings Then
13            RowNumbersOrHeaders = sArrayIf(sArrayIsText(RowNumbersOrHeaders), sMatch(RowNumbersOrHeaders, sSubArray(TheArray, 1, 1, , 1)), RowNumbersOrHeaders)
14        End If
15        For Each x In ColumnNumbersOrHeaders
16            If VarType(x) = vbString Then
17                CHasStrings = True
18                Exit For
19            End If
20        Next
21        If CHasStrings Then
22            ColumnNumbersOrHeaders = sArrayIf(sArrayIsText(ColumnNumbersOrHeaders), sMatch(ColumnNumbersOrHeaders, sArrayTranspose(sSubArray(TheArray, 1, 1, 1))), ColumnNumbersOrHeaders)
23        End If

24        On Error GoTo ErrHandler
25        NR = 1: NC = 1
26        If True Then
27            R = sNRows(RowNumbersOrHeaders): c = sNCols(RowNumbersOrHeaders)
28            If R = 1 Then RowLock(1) = True
29            If c = 1 Then ColLock(1) = True
30            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
31            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
32        End If
33        If True Then
34            R = sNRows(ColumnNumbersOrHeaders): c = sNCols(ColumnNumbersOrHeaders)
35            If R = 1 Then RowLock(2) = True
36            If c = 1 Then ColLock(2) = True
37            If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
38            If c > 1 Then If (NC = 1 Or c < NC) Then NC = c
39        End If

40        ReDim Result(1 To NR, 1 To NC)

41        For i = 1 To NR
42            For j = 1 To NC
43                Result(i, j) = sIndexSafe(TheArray, _
                      RowNumbersOrHeaders(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j)), _
                      ColumnNumbersOrHeaders(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j)))
44            Next j
45        Next i
46        sIndex = Result

47        Exit Function
ErrHandler:
48        sIndex = "#sIndex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function sIndexSafe(ByRef TheArray, i, j)
1         On Error GoTo ErrHandler
2         If i <> CLng(i) Then
3             sIndexSafe = "#Illegal index!"
4         ElseIf j <> CLng(j) Then
5             sIndexSafe = "#Illegal index!"
6         Else
7             sIndexSafe = TheArray(i, j)
8         End If
9         Exit Function
ErrHandler:
10        sIndexSafe = "#Illegal index!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIntegers
' Author    : Philip Swannell
' Date      : 20-Jun-2013
' Purpose   : Returns a column array of integers from 1 to N.
' Arguments
' N         : A positive integer
' -----------------------------------------------------------------------------------------------------------------------
Function sIntegers(N As Long)
Attribute sIntegers.VB_Description = "Returns a column array of integers from 1 to N."
Attribute sIntegers.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim TheReturn() As Long
1         On Error GoTo ErrHandler

2         If N <= 0 Then sIntegers = "#N must be positive!"
3         ReDim TheReturn(1 To N, 1 To 1)
4         For i = 1 To N
5             TheReturn(i, 1) = i
6         Next i
7         sIntegers = TheReturn

8         Exit Function
ErrHandler:
9         sIntegers = "#sIntegers (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIsRegMatch
' Author    : Philip Swannell
' Date      : 07-Dec-2015
' Purpose   : Implements Regular Expressions exposed by "Microsoft VBScript Regular Expressions 5.5".
'             The function returns TRUE if StringToSearch matches RegularExpression, FALSE
'             if it does not match, or an error string if RegularExpression contains a
'             syntax error.
' Arguments
' RegularExpression: The regular expression. Must be a string. Example cat|dog to match on either the string
'             cat or the string dog.
' StringToSearch: The string to match. May be an array in which case the return from the function is an
'             array of the same dimensions.
' CaseSensitive: TRUE for case-sensitive matching, FALSE for case-insensitive matching. This argument is
'             optional, defaulting to FALSE for case-insensitive matching.
'
' Notes     : Syntax cheat sheet:
'             Character classes
'             .                 any character except newline
'             \w \d \s          word, digit, whitespace
'             \W \D \S          not word, not digit, not whitespace
'             [abc]             any of a, b, or c
'             [^abc]            not a, b, or c
'             [a-g]             character between a & g
'
'             Anchors
'             ^abc$              start / end of the string
'             \b                 word boundary
'
'             Escaped characters
'             \. \* \\          escaped special characters
'             \t \n \r          tab, linefeed, carriage return
'
'             Groups and Look-arounds
'             (abc)             capture group
'             \1                backreference to group #1
'             (?:abc)           non-capturing group
'             (?=abc)           positive lookahead
'             (?!abc)           negative lookahead
'
'             Quantifiers and Alternation
'             a* a+ a?          0 or more, 1 or more, 0 or 1
'             a{5} a{2,}        exactly five, two or more
'             a{1,3}            between one & three
'             a+? a{2,}?        match as few as possible
'             ab|cd             match ab or cd
'
'             Further reading:
'             http://www.regular-expressions.info/
'             https://en.wikipedia.org/wiki/Regular_expression
' -----------------------------------------------------------------------------------------------------------------------
Function sIsRegMatch(RegularExpression As String, ByVal StringToSearch As Variant, Optional CaseSensitive As Boolean = False)
Attribute sIsRegMatch.VB_Description = "Implements Regular Expressions exposed by ""Microsoft VBScript Regular Expressions 5.5"". The function returns TRUE if StringToSearch matches RegularExpression, FALSE if it does not match, or an error string if RegularExpression contains a syntax error."
Attribute sIsRegMatch.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim Result() As Variant
          Dim rx As VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             sIsRegMatch = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If
6         Set rx = New RegExp
7         With rx
8             .IgnoreCase = Not CaseSensitive
9             .Pattern = RegularExpression
10            .Global = False        'Find first match only
11        End With

12        If VarType(StringToSearch) = vbString Then
13            sIsRegMatch = rx.Test(StringToSearch)

14            GoTo EarlyExit
15        ElseIf VarType(StringToSearch) < vbArray Then
16            sIsRegMatch = "#StringToSearch must be a string!"
17            GoTo EarlyExit
18        End If
19        If TypeName(StringToSearch) = "Range" Then StringToSearch = StringToSearch.Value2

20        Select Case NumDimensions(StringToSearch)
              Case 2
21                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1), LBound(StringToSearch, 2) To UBound(StringToSearch, 2))
22                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
23                    For j = LBound(StringToSearch, 2) To UBound(StringToSearch, 2)
24                        If VarType(StringToSearch(i, j)) = vbString Then
25                            Result(i, j) = rx.Test(StringToSearch(i, j))
26                        Else
27                            Result(i, j) = "#StringToSearch must be a string!"
28                        End If
29                    Next j
30                Next i
31            Case 1
32                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1))
33                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
34                    If VarType(StringToSearch(i)) = vbString Then
35                        Result(i) = rx.Test(StringToSearch(i))
36                    Else
37                        Result(i) = "#StringToSearch must be a string!"
38                    End If
39                Next i
40            Case Else
41                Throw "StringToSearch must be String or array with 1 or 2 dimensions"
42        End Select

43        sIsRegMatch = Result
EarlyExit:
44        Set rx = Nothing

45        Exit Function
ErrHandler:
46        sIsRegMatch = "#sIsRegMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
47        Set rx = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatch
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : Return the integer row number of the value(s) LookupValues in the column LookupColumn. If
'             the value cannot be found, an error is returned. Row counting starts at 1.
'             Like Excel MATCH with match_type of 0, but allows array LookupValues when
'             called from VBA.
' Arguments
' LookUpValues: An array of values to look up. Passing an array into an array formula is likely to be
'             faster than many formulas each handling a single value.
' LookupColumn: A column of values to search in. May be passed in any order unless LookupColumnIsSorted is
'             TRUE, in which case it must be sorted in ascending order.
' LookupColumnIsSorted: If LookupColumn is known to be already sorted in ascending order then passing as TRUE will
'             provide a speed-up for large array sizes. If TRUE but LookupColumn is not
'             sorted in ascending the function returns an error.
'
' Notes     : The function is similar to Excel's MATCH with argument match_type set to zero, but has the
'             advantage that when called from VBA the first argument may be passed as an
'             array. Implementation is via one of two algorithms according to which is
'             likely to be faster, i.e. either a) repeated linear search; or b) sorting of
'             LookupColumn followed by repeated binary chop search.
' -----------------------------------------------------------------------------------------------------------------------
Function sMatch(ByVal LookupValues As Variant, ByVal LookupColumn, Optional LookupColumnIsSorted As Boolean)
Attribute sMatch.VB_Description = "Return the integer row number of the value(s) LookupValues in the column LookupColumn. If the value cannot be found, an error is returned. Row counting starts at 1. Like Excel MATCH with match_type of 0, but allows array LookupValues when called from VBA."
Attribute sMatch.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim M As Long
          Dim N As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti LookupValues, LookupColumn

          'Cope with multiple columns in LookUpValues by reshaping to single column and reshaping back _
           before function exit- would be better to change each of the functions sMatchNaive, sMatchNaive2 _
           and sMatchSortChop _
           TODO fix this!
3         If sNCols(LookupValues) > 1 Then
              Dim didReshape As Boolean
              Dim origNC As Long
              Dim origNR As Long
4             didReshape = True
5             origNR = sNRows(LookupValues)
6             origNC = sNCols(LookupValues)
7             LookupValues = sReshape(LookupValues, origNR * origNC, 1)
8         End If

9         If sNCols(LookupColumn) > 1 Then Throw "LookupColumn cannot be a multi-column array. Use sMultiMatch instead."

10        N = sNRows(LookupValues): M = sNRows(LookupColumn)

11        If LookupColumnIsSorted Then
12            sMatch = sMatchSortChop(LookupValues, LookupColumn, N, M, True)
13        Else
14            Select Case sMatchWhichAlgorithm(N, M)
                  Case 1
15                    sMatch = sMatchNaive(LookupValues, LookupColumn, N, M)
16                Case 2
17                    sMatch = sMatchNaive2(LookupValues, LookupColumn, N)
18                Case 3
19                    sMatch = sMatchSortChop(LookupValues, LookupColumn, N, M, LookupColumnIsSorted)
20            End Select
21        End If

22        If didReshape Then
23            sMatch = sReshape(sMatch, origNR, origNC)
24        End If

25        Exit Function
ErrHandler:
26        sMatch = "#sMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIsIn
' Author    : Philip Swannell
' Date      : 13-Feb-2020
' Purpose   : Determine if each element of Items is an element of TheArray.
' Arguments
' Items     : An array of arbitrary values. The dimensions of the return matches the dimensions of this
'             argument.
' TheArray  : An array of arbitrary values.
' TheArrayIsSorted: Defaults to False, in which case TheArray does not need to be sorted. If TheArray is
'             sorted then pass True for better speed. Sorted = increasing from top to
'             bottom (a 1-column array),  left to right (1 row) or "like a book" (>1 row,
'             >1 column).
' -----------------------------------------------------------------------------------------------------------------------
Function sIsIn(ByVal Items As Variant, ByVal TheArray, Optional TheArrayIsSorted As Boolean)
Attribute sIsIn.VB_Description = "Determine if each element of Items is an element of TheArray."
Attribute sIsIn.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim NR As Long, NC As Long
1         Force2DArrayR TheArray, NR, NC
2         If NC > 1 Then
3             If NR > 1 Then
4                 TheArray = sReshape(TheArray, , 1)
5             End If
6         End If
7         sIsIn = sArrayIsNumber(sMatch(Items, TheArray, TheArrayIsSorted))
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatchInMatrix
' Author    : Philip Swannell
' Date      : 29-Mar-2018
' Purpose   : Returns row and column numbers of the value(s) LookupValues in the matrix LookupMatrix. If
'             the value is not found, an error is returned. Row and column counting starts
'             at 1. Return is a string such as '2,3' if the value is found at row 2, column
'             3.
' Arguments
' LookupValues: An arbitrary value. May be an array
' LookupMatrix: An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sMatchInMatrix(ByVal LookupValues, ByVal LookupMatrix)
Attribute sMatchInMatrix.VB_Description = "Returns row and column numbers of the value(s) LookupValues in the matrix LookupMatrix. If the value is not found, an error is returned. Row and column counting starts at 1. Return is a string such as '2,3' if the value is found at row 2, column 3."
Attribute sMatchInMatrix.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim c As Long
          Dim i As Long
          Dim j As Long
          Dim MatchRes
          Dim NC As Long
          Dim NC2 As Long
          Dim NR As Long
          Dim NR2 As Long
          Dim R As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti LookupValues, LookupMatrix

3         NR = sNRows(LookupMatrix)
4         NC = sNCols(LookupMatrix)

5         MatchRes = sMatch(LookupValues, sReshape(LookupMatrix, NR * NC, 1))
6         If IsArray(MatchRes) Then
7             NR2 = sNRows(MatchRes)
8             NC2 = sNCols(MatchRes)
9             For i = 1 To NR2
10                For j = 1 To NC2
11                    If IsNumber(MatchRes(i, j)) Then
12                        c = (MatchRes(i, j) - 1) Mod NC + 1
13                        R = (MatchRes(i, j) - c) \ NC + 1
14                        MatchRes(i, j) = CStr(R) + "," + CStr(c)
15                    End If
16                Next j
17            Next i
18        ElseIf IsNumber(MatchRes) Then
19            c = (MatchRes - 1) Mod NC + 1
20            R = (MatchRes - c) \ NC + 1
21            MatchRes = CStr(R) + "," + CStr(c)
22        End If
23        sMatchInMatrix = MatchRes
24        Exit Function
ErrHandler:
25        sMatchInMatrix = "#sMatchInMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatchNaive
' Author    : Philip Swannell
' Date      : 19-Apr-2015
' Purpose   : Naive algorithm - performance of order NM, but faster when either N or M is sufficiently small
'             N = number of rows in lookupValues, M = number of rows in LookupColumn
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMatchNaive(LookupValues, LookupColumn, N As Long, M As Long)
          Dim i As Long
          Dim j As Long
          Dim Result As Variant

1         On Error GoTo ErrHandler

2         Result = sReshape(0, N, 1)
3         For i = 1 To N
4             For j = 1 To M
5                 If sEquals(LookupValues(i, 1), LookupColumn(j, 1)) Then
6                     Result(i, 1) = j
7                     Exit For
8                 End If
9             Next j
10            If Result(i, 1) = 0 Then Result(i, 1) = "#Element not found!"
11        Next i
12        If N = 1 Then
13            sMatchNaive = Result(1, 1)
14        Else
15            sMatchNaive = Result
16        End If

17        Exit Function
ErrHandler:
18        Throw "#sMatchNaive (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatchNaive2
' Author    : Philip Swannell
' Date      : 19-Apr-2015
' Purpose   : Wrap of Worksheet function MATCH. Wrap necessary since that function cannot
'             accept an array as first argument when called from VBA - though it can when
'             called from the sheet :-(
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMatchNaive2(LookupValues, LookupColumn, N As Long)
          Dim i As Long
          Dim Result As Variant

1         On Error GoTo ErrHandler

2         Result = sReshape(0, N, 1)
3         For i = 1 To N
4             On Error Resume Next
5             Result(i, 1) = Application.WorksheetFunction.Match(LookupValues(i, 1), LookupColumn, 0)
6             On Error GoTo ErrHandler

7             If Result(i, 1) = 0 Then Result(i, 1) = "#Element not found!"
8         Next i
9         If N = 1 Then
10            sMatchNaive2 = Result(1, 1)
11        Else
12            sMatchNaive2 = Result
13        End If

14        Exit Function
ErrHandler:
15        Throw "#sMatchNaive2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatchSortChop
' Author    : Philip Swannell
' Date      : 19-Apr-2015
' Purpose   : Sub routine of sMatch, used when neither N (# of elements in LookupValues) or
'             M (# in LookupColumn) is small. Algorithm is to first sort the LookupColumn, then
'             remove duplicates then search via binary chop.
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMatchSortChop(LookupValues, LookupColumn, N As Long, M As Long, AlreadySorted As Boolean)
          Dim Bottom As Long
          Dim ChooseVector As Variant
          Dim ColumnWithIndex As Variant
          Dim Found As Boolean
          Dim i As Long
          Dim Middle As Long
          Dim repeatsExist As Boolean
          Dim Result As Variant
          Dim Top As Long

1         On Error GoTo ErrHandler

2         ColumnWithIndex = sArrayRange(LookupColumn, sIntegers(M))
3         If Not AlreadySorted Then
              'PGS 24/4/2019 Important to pass argument UseExcelSortMethod as False below, since comparison within the _
               chop search is done via method VariantLessThan, which is exactly same comparison method _
               used in sSortedArray if UseExcelSortMethod is FALSE. Otherwise subtle bugs can occur with _
               bogus return of "#Element not found!"
4             ColumnWithIndex = sSortedArray(ColumnWithIndex, , , , , , , , False)
5         Else
6             For i = 2 To M
7                 If VariantLessThan(LookupColumn(i, 1), LookupColumn(i - 1, 1), False) Then Throw "LookupColumn is not sorted"
8             Next i
9         End If

10        ChooseVector = sReshape(0, M, 1)
11        ChooseVector(1, 1) = True
12        For i = 2 To M
13            If sEquals(ColumnWithIndex(i, 1), ColumnWithIndex(i - 1, 1)) Then
14                repeatsExist = True
15                ChooseVector(i, 1) = False
16            Else
17                ChooseVector(i, 1) = True
18            End If
19        Next i

20        If repeatsExist Then
21            ColumnWithIndex = sMChoose(ColumnWithIndex, ChooseVector)
22            M = sNRows(ColumnWithIndex)
23        End If

24        Result = sReshape(0, N, 1)

25        For i = 1 To N
26            Found = False
27            Top = 1: Bottom = M
28            Do While Bottom - Top > 1
29                Middle = (Top + Bottom) / 2
30                If VariantLessThan(ColumnWithIndex(Middle, 1), LookupValues(i, 1), False) Then
31                    Top = Middle
32                ElseIf sEquals(ColumnWithIndex(Middle, 1), LookupValues(i, 1), False) Then
33                    Found = True
34                    Top = Middle: Bottom = Middle
35                Else
36                    Bottom = Middle
37                End If
38            Loop
39            If Found Then
40                Result(i, 1) = ColumnWithIndex(Middle, 2)
41            ElseIf sEquals(ColumnWithIndex(Top, 1), LookupValues(i, 1)) Then
42                Result(i, 1) = ColumnWithIndex(Top, 2)
43            ElseIf sEquals(ColumnWithIndex(Bottom, 1), LookupValues(i, 1)) Then
44                Result(i, 1) = ColumnWithIndex(Bottom, 2)
45            Else
46                Result(i, 1) = "#Element not found!"
47            End If
48        Next i
49        If N = 1 Then
50            sMatchSortChop = Result(1, 1)
51        Else
52            sMatchSortChop = Result
53        End If
54        Exit Function
ErrHandler:
55        Throw "#sMatchSortChop (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMatchWhichAlgorithm
' Author    : Philip
' Date      : 31-Dec-2015
' Purpose   : Encapsulate the choice of algorithm for sMatch
'             See https://d.docs.live.net/4251b448d4115355/Excel Sheets/Match Algorithm test v2.xlsm
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMatchWhichAlgorithm(N As Long, M As Long) As Long
          Dim TimeNaive As Double
          Dim TimeNaive2 As Double
          Dim TimeSortChop As Double

1         On Error GoTo ErrHandler

2         TimeNaive = 3.85517869441564E-03 + 7.23472188814331E-04 * N + -3.03004938537298E-05 * M + 2.25868273197574E-04 * N * M
3         TimeNaive2 = 3.6156827209229E-03 + 4.15102223197502E-03 * N + 1.73374818312023E-05 * M + 6.67576107110462E-05 * N * M
4         TimeSortChop = 1.23347244974755E-02 + 1.30243494210562E-03 * N + 2.98433751641394E-03 * M + 4.04082409890375E-07 * N * M

5         If (TimeNaive < TimeSortChop) And (TimeNaive < TimeNaive2) Then
6             sMatchWhichAlgorithm = 1        'sMatchNaive
7         ElseIf (TimeNaive2 < TimeSortChop) And (TimeNaive2 < TimeNaive) Then
8             sMatchWhichAlgorithm = 2        'sMatchNaive2
9         Else
10            sMatchWhichAlgorithm = 3        'sMatchSortChop
11        End If
12        Exit Function
ErrHandler:
13        sMatchWhichAlgorithm = "#sMatchWhichAlgorithm (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : sSearchSorted
' Author    : Philip Swannell
' Date      : 10-Feb-2022
' Purpose   : Returns the index in LookupColumn of each element of LookupValues.
' Arguments
' LookupValues: A value or a 1-column array of values.
' LookupColumn: A 1-column array of values, which must be sorted in ascending order i.e. with smaller
'             elements above larger elements.
' IntermediatesMatchToLarger: If TRUE then when a LookupValue lies between adjacent elements of LookupColumn the
'             function returns the index of the larger of the two, otherwise it returns the
'             index of the smaller.
'---------------------------------------------------------------------------------------------------------
Function sSearchSorted(ByVal LookupValues, ByVal LookupColumn, IntermediatesMatchToLarger As Boolean)
Attribute sSearchSorted.VB_Description = "Returns the index in LookupColumn of each element of LookupValues."
Attribute sSearchSorted.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim Bottom As Long
          Dim ChooseVector As Variant
          Dim ColumnWithIndex As Variant
          Dim Found As Boolean
          Dim i As Long
          Dim Middle As Long
          Dim repeatsExist As Boolean
          Dim Result As Variant
          Dim Top As Long
          Dim NC As Long
          Dim NumValues As Long
          Dim NumInColumn As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR LookupValues, NumValues, NC
3         If NC <> 1 Then Throw "LookupValues must have a single column"
4         Force2DArrayR LookupColumn, NumInColumn, NC
5         If NC <> 1 Then Throw "LookupColumn must have a single column"

6         ColumnWithIndex = sArrayRange(LookupColumn, sIntegers(NumInColumn))
7         For i = 2 To NumInColumn
8             If VariantLessThan(LookupColumn(i, 1), LookupColumn(i - 1, 1), False) Then Throw "LookupColumn is not sorted. Element " & CStr(i - 1) & " exceeds element " & CStr(i)
9         Next i

10        ChooseVector = sReshape(0, NumInColumn, 1)
11        ChooseVector(1, 1) = True
12        For i = 2 To NumInColumn
13            If sEquals(ColumnWithIndex(i, 1), ColumnWithIndex(i - 1, 1)) Then
14                repeatsExist = True
15                ChooseVector(i, 1) = False
16            Else
17                ChooseVector(i, 1) = True
18            End If
19        Next i

20        If repeatsExist Then
21            ColumnWithIndex = sMChoose(ColumnWithIndex, ChooseVector)
22            NumInColumn = sNRows(ColumnWithIndex)
23        End If

24        Result = sReshape(0, NumValues, 1)

25        For i = 1 To NumValues
26            Found = False
27            Top = 1: Bottom = NumInColumn
28            Do While Bottom - Top > 1
29                Middle = (Top + Bottom) / 2
30                If VariantLessThan(ColumnWithIndex(Middle, 1), LookupValues(i, 1), False) Then
31                    Top = Middle
32                ElseIf sEquals(ColumnWithIndex(Middle, 1), LookupValues(i, 1), False) Then
33                    Found = True
34                    Top = Middle: Bottom = Middle
35                Else
36                    Bottom = Middle
37                End If
38            Loop
39            If Found Then
40                Result(i, 1) = ColumnWithIndex(Middle, 2)
41            ElseIf sEquals(ColumnWithIndex(Top, 1), LookupValues(i, 1)) Then
42                Result(i, 1) = ColumnWithIndex(Top, 2)
43            ElseIf sEquals(ColumnWithIndex(Bottom, 1), LookupValues(i, 1)) Then
44                Result(i, 1) = ColumnWithIndex(Bottom, 2)
45            Else
                  'No exact match
46                If VariantLessThan(LookupValues(i, 1), ColumnWithIndex(1, 1), False) Then
                      'Before the first
47                    Result(i, 1) = IIf(IntermediatesMatchToLarger, 1, 0)
48                ElseIf VariantLessThan(ColumnWithIndex(NumInColumn, 1), LookupValues(i, 1), False) Then
                      'after the last
49                    Result(i, 1) = ColumnWithIndex(NumInColumn, 2) + IIf(IntermediatesMatchToLarger, 1, 0)
50                Else
51                    Result(i, 1) = ColumnWithIndex(IIf(IntermediatesMatchToLarger, Bottom, Top), 2)
52                End If
53            End If
54        Next i
55        If NumValues = 1 Then
56            sSearchSorted = Result(1, 1)
57        Else
58            sSearchSorted = Result
59        End If
60        Exit Function
ErrHandler:
61        sSearchSorted = "#sSearchSorted (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMaxOfNums
' Author    : Philip Swannell
' Date      : 15-Sep-2015
' Purpose   : Returns the maximum of the numerical elements of the input TheArray, ignoring those
'             elements which are not numeric. If no elements are numeric, an error is
'             returned. Note that strings containing numbers, such as "123", will be
'             ignored.
' Arguments
' TheArray  : Input array of elements of any type. Elements that are not numbers will be ignored.
' -----------------------------------------------------------------------------------------------------------------------
Function sMaxOfNums(TheArray As Variant)
Attribute sMaxOfNums.VB_Description = "Returns the maximum of the numerical elements of the input TheArray, ignoring those elements which are not numeric. If no elements are numeric, an error is returned. Note that strings containing numbers, such as ""123"", will be ignored."
Attribute sMaxOfNums.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
          Dim c As Variant
          Dim Res As Variant

3         For Each c In TheArray
4             If IsNumberOrDate(c) Then
5                 If IsEmpty(Res) Then
6                     Res = c
7                 ElseIf c > Res Then
8                     Res = c
9                 End If
10            End If
11        Next
12        If IsEmpty(Res) Then
13            sMaxOfNums = "#No numbers found in TheArray!"
14        Else
15            sMaxOfNums = Res
16        End If

17        Exit Function
ErrHandler:
18        sMaxOfNums = "#sMaxOfNums (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMinOfNums
' Author    : Philip Swannell
' Date      : 15-Sep-2015
' Purpose   : Returns the minimum of the numerical elements of the input TheArray, ignoring those
'             elements which are not numeric. If no elements are numeric, an error is
'             returned. Note that strings containing numbers, such as "123", will be
'             ignored.
' Arguments
' TheArray  : Input array of elements of any type. Elements that are not numbers will be ignored.
' -----------------------------------------------------------------------------------------------------------------------
Function sMinOfNums(TheArray As Variant)
Attribute sMinOfNums.VB_Description = "Returns the minimum of the numerical elements of the input TheArray, ignoring those elements which are not numeric. If no elements are numeric, an error is returned. Note that strings containing numbers, such as ""123"", will be ignored."
Attribute sMinOfNums.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
          Dim c As Variant
          Dim Res As Variant

3         For Each c In TheArray
4             If IsNumberOrDate(c) Then
5                 If IsEmpty(Res) Then
6                     Res = c
7                 ElseIf c < Res Then
8                     Res = c
9                 End If
10            End If
11        Next
12        If IsEmpty(Res) Then
13            sMinOfNums = "#No numbers found in TheArray!"
14        Else
15            sMinOfNums = Res
16        End If

17        Exit Function
ErrHandler:
18        sMinOfNums = "#sMinOfNums (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMultiMatch
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Matches entire rows of LookupValues in the rows of LookupArray. If a row cannot be found ,
'             an error is returned. Row counting starts at 1.
' Arguments
' LookUpValues: An array of arbitrary values.
' LookupArray: An array of arbitrary values. Must have the same number of columns as LookupValues.
' CaseSensitive: Whether the matching of strings should be sensitive to case.
' -----------------------------------------------------------------------------------------------------------------------
Function sMultiMatch(LookupValues, LookupArray, CaseSensitive As Boolean)
Attribute sMultiMatch.VB_Description = "Matches entire rows of LookupValues in the rows of LookupArray. If a row cannot be found , an error is returned. Row counting starts at 1."
Attribute sMultiMatch.VB_ProcData.VB_Invoke_Func = " \n27"

1         On Error GoTo ErrHandler
          Dim M As Long
          Dim N As Long
          Dim NC As Long
          Dim TimeNaive As Double
          Dim TimeSortChop As Double

2         Force2DArrayRMulti LookupValues, LookupArray

3         N = sNRows(LookupValues)
4         M = sNRows(LookupArray)

          'Estimate times (miliseconds) for each of the algorithms. We will use Naive when either N or M is "small"
          'Workbook used to estimate parameters on Philip Swannell's home PC
          'C:\Users\Philip\OneDrive\Excel Sheets\Match Algorithm test.xlsm
          'or this link should work from elsewhere:
          'https://onedrive.live.com/redir?page=view&resid=4251B448D4115355!24821&authkey=!AF9L_242FkPW-Ag

5         TimeNaive = 0.016214 + 0.0006092 * N + 0 * M + 0.00020806 * N * M
6         TimeSortChop = 0.014371 + 0.0013418 * N + 0.0097363 * M + 0.00000067149 * M * N

7         NC = sNCols(LookupValues)

8         If NC <> sNCols(LookupArray) Then Throw "LookupValues and LookupArray must have the same number of columns"

9         If TimeNaive < TimeSortChop Then
10            sMultiMatch = sMultiMatchNaive(LookupValues, LookupArray, CaseSensitive)
11        Else
12            sMultiMatch = sMultiMatchSortChop(LookupValues, LookupArray, N, M, NC, CaseSensitive, False)
13        End If

14        Exit Function
ErrHandler:
15        sMultiMatch = "#sMultiMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMultiMatchNaive
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Naive algorithm - performance of order NM, but faster when either N or M is sufficiently small
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMultiMatchNaive(LookupValues, LookupArray, CaseSensitive As Boolean)
          Dim Found As Boolean
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NRA As Long
          Dim Result() As Variant
1         On Error GoTo ErrHandler

2         NC = sNCols(LookupValues)
3         If NC <> sNCols(LookupArray) Then Throw "LookupValues and LookupArray must have the same number of columns"
4         NR = sNRows(LookupValues)
5         NRA = sNRows(LookupArray)
6         ReDim Result(1 To NR, 1 To 1)

7         For i = 1 To NR
8             Found = False
9             For j = 1 To NRA
10                If sMultiMatchRowsEqual(LookupValues, i, LookupArray, j, NC, CaseSensitive) Then
11                    Found = True
12                    Result(i, 1) = j
13                    Exit For
14                End If
15            Next j
16            If Not Found Then Result(i, 1) = "#Row not found!"
17        Next i
18        If NR = 1 Then
19            sMultiMatchNaive = Result(1, 1)
20        Else
21            sMultiMatchNaive = Result
22        End If

23        Exit Function
ErrHandler:
24        sMultiMatchNaive = "#sMultiMatchNaive (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMultiMatchRowLessThan
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Is row i of Array1 less than row j of Array2?
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMultiMatchRowLessThan(Array1, i As Long, Array2, j As Long, NC As Long, CaseSensitive As Boolean) As Boolean
          Dim k As Long

1         On Error GoTo ErrHandler
2         For k = 1 To NC
3             If VariantLessThan(Array1(i, k), Array2(j, k), CaseSensitive) Then
4                 sMultiMatchRowLessThan = True
5                 Exit Function
6             ElseIf Not sEquals(Array1(i, k), Array2(j, k), CaseSensitive) Then
7                 sMultiMatchRowLessThan = False
8                 Exit Function
9             End If
10        Next k
11        sMultiMatchRowLessThan = False        'the two rows are equal, but that's false as this is a LessThan function, not a LessThanOrEquals function
12        Exit Function
ErrHandler:
13        Throw "#sMultiMatchRowLessThan (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMultiMatchRowsEqual
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Is the ith row of Array1 the same as the jth row of Array2
' -----------------------------------------------------------------------------------------------------------------------
Private Function sMultiMatchRowsEqual(Array1, i As Long, Array2, j As Long, NC As Long, CaseSensitive As Boolean) As Boolean
          Dim k As Long

1         On Error GoTo ErrHandler
2         For k = 1 To NC
3             If Not sEquals(Array1(i, k), Array2(j, k), CaseSensitive) Then
4                 sMultiMatchRowsEqual = False
5                 Exit Function
6             End If
7         Next k

8         sMultiMatchRowsEqual = True
9         Exit Function
ErrHandler:
10        Throw "#sMultiMatchRowsEqual (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMultiMatchSortChop
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : sort-chop version of MultiMatch
'             argument alreadySorted means already sorted and duplicates removed!
' -----------------------------------------------------------------------------------------------------------------------
Function sMultiMatchSortChop(LookupValues, LookupArray, NRV As Long, NRA As Long, NC As Long, CaseSensitive As Boolean, AlreadySorted As Boolean)
          Dim ArrayWithIndex As Variant
          Dim Bottom As Long
          Dim ChooseVector As Variant
          Dim Found As Boolean
          Dim i As Long
          Dim Middle As Long
          Dim repeatsExist As Boolean
          Dim Result As Variant
          Dim Top As Long

1         On Error GoTo ErrHandler

2         If Not AlreadySorted Then
3             ArrayWithIndex = sArrayRange(LookupArray, sIntegers(NRA))
4             ArrayWithIndex = sSortedArray2(ArrayWithIndex, sIntegers(NC), sReshape(True, NC, 1), CaseSensitive)

5             ChooseVector = sReshape(0, NRA, 1)
6             ChooseVector(1, 1) = True
7             For i = 2 To NRA
8                 If sMultiMatchRowsEqual(ArrayWithIndex, i, ArrayWithIndex, i - 1, NC, CaseSensitive) Then
9                     repeatsExist = True
10                    ChooseVector(i, 1) = False
11                Else
12                    ChooseVector(i, 1) = True
13                End If
14            Next i

15            If repeatsExist Then
16                ArrayWithIndex = sMChoose(ArrayWithIndex, ChooseVector)
17                NRA = sNRows(ArrayWithIndex)
18            End If
19        Else
20            ArrayWithIndex = LookupArray        'It's called ArrayWithIndex but in this case it doesn't have an index
21        End If

22        Result = sReshape(0, NRV, 1)

23        For i = 1 To NRV
24            Found = False
25            Top = 1: Bottom = NRA
26            Do While Bottom - Top > 1
27                Middle = (Top + Bottom) / 2
28                If sMultiMatchRowLessThan(ArrayWithIndex, Middle, LookupValues, i, NC, CaseSensitive) Then
29                    Top = Middle
30                ElseIf sMultiMatchRowsEqual(ArrayWithIndex, Middle, LookupValues, i, NC, CaseSensitive) Then
31                    Found = True
32                    Top = Middle: Bottom = Middle
33                Else
34                    Bottom = Middle
35                End If
36            Loop
37            If Found Then
38                If AlreadySorted Then
39                    Result(i, 1) = Middle
40                Else
41                    Result(i, 1) = ArrayWithIndex(Middle, NC + 1)
42                End If
43            ElseIf sMultiMatchRowsEqual(ArrayWithIndex, Top, LookupValues, i, NC, CaseSensitive) Then
44                If AlreadySorted Then
45                    Result(i, 1) = Top
46                Else
47                    Result(i, 1) = ArrayWithIndex(Top, NC + 1)
48                End If
49            ElseIf sMultiMatchRowsEqual(ArrayWithIndex, Bottom, LookupValues, i, NC, CaseSensitive) Then
50                If AlreadySorted Then
51                    Result(i, 1) = Bottom
52                Else
53                    Result(i, 1) = ArrayWithIndex(Bottom, NC + 1)
54                End If
55            Else
56                Result(i, 1) = "#Row not found!"
57            End If
58        Next i
59        If NRV = 1 Then
60            sMultiMatchSortChop = Result(1, 1)
61        Else
62            sMultiMatchSortChop = Result
63        End If
64        Exit Function
ErrHandler:
65        Throw "#sMultiMatchSortChop (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sInsert
' Author    : Philip Swannell
' Date      : 23-Jul-2019
' Purpose   : Inserts TheValue into TheArray at position (RowNum,ColNum). TheValue, RowNum and ColNum
'             may be arrays.
' Arguments
' TheArray  : Array of arbitrary values.
' TheValue  : The value to be inserted into TheArray. May be an array of values.
' RowNum    : The row number of TheArray at which TheValue is inserted. May be a 1-column array of row
'             numbers.
' ColNum    : The column number of TheArray at which TheValue is inserted. May be a 1-row array of
'             column numbers.
' Operation : Omit or 'Overwrite' to overwrite the existing value within TheArray with TheValue. Also
'             supported: 'Add', 'Subtract', 'Multiply' and 'Divide', 'Concatenate'.
'
' Notes     : Example:
'             If all cells in the range A1:C3 are zero then the formula
'             =sInsert(A1:D3, "X", {1;3}, {1,4})
'             returns the array:
'             X          0          0          X
'             0          0          0          0
'             X          0          0          X
' -----------------------------------------------------------------------------------------------------------------------
Function sInsert(ByVal TheArray, TheValue, Optional ByVal RowNum, Optional ByVal ColNum, Optional Operation As String = "Overwrite")
Attribute sInsert.VB_Description = "Inserts TheValue into TheArray at position (RowNum,ColNum). TheValue, RowNum and ColNum may be arrays."
Attribute sInsert.VB_ProcData.VB_Invoke_Func = " \n24"

          Dim NC_CN As Long
          Dim NC_RN As Long
          Dim NC_TA As Long
          Dim NC_TV As Long
          Dim NR_CN As Long
          Dim NR_RN As Long
          Dim NR_TA As Long
          Dim NR_TV As Long
          Const RowNumErr = "RowNum must be an integer between 1 and the number of rows in TheArray, or a 1-column array of such. Omit to specify all rows"
          Const ColNumErr = "ColNum must be an integer between 1 and the number of colums in TheArray, or a 1-row array of such. Omit to specify all columns"
          Dim i As Long
          Dim j As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, NR_TA, NC_TA
3         If IsMissing(RowNum) Then RowNum = sIntegers(NR_TA)
4         If IsMissing(ColNum) Then ColNum = sArrayTranspose(sIntegers(NC_TA))

5         Force2DArrayR TheValue, NR_TV, NC_TV
6         Force2DArrayR RowNum, NR_RN, NC_RN
7         Force2DArrayR ColNum, NR_CN, NC_CN

          'Validate RowNum
8         If NC_RN > 1 Then Throw RowNumErr + ", but it has more than one column"
9         For i = 1 To NR_RN
10            If Not IsNumber(RowNum(i, 1)) Then Throw RowNumErr + ", but element " + CStr(i) + " is not a number"
11            If RowNum(i, 1) < 1 Then Throw RowNumErr + ", but element " + CStr(i) + " less than 1"
12            If RowNum(i, 1) > NR_TA Then Throw RowNumErr + ", but element " + CStr(i) + " is greater than the number of rows in TheArray"
13            If CLng(RowNum(i, 1)) <> RowNum(i, 1) Then Throw RowNumErr + ", but element " + CStr(i) + " is not a whole number"
14        Next i

          'Validate ColNum
15        If NR_CN > 1 Then Throw ColNumErr + ", but it has more than one row"
16        For i = 1 To NC_CN
17            If Not IsNumber(ColNum(1, i)) Then Throw ColNumErr + ", but element " + CStr(i) + " is not a number"
18            If ColNum(1, i) < 1 Then Throw ColNumErr + ", but element " + CStr(i) + " less than 1"
19            If ColNum(1, i) > NC_TA Then Throw ColNumErr + ", but element " + CStr(i) + " is greater than the number of columns in TheArray"
20            If CLng(ColNum(1, i)) <> ColNum(1, i) Then Throw ColNumErr + ", but element " + CStr(i) + " is not a whole number"
21        Next i

          Dim WriteColIncrements As Boolean
          Dim WriteRowIncrements As Boolean

          'Handle case when TheValue as more than 1 row, and RowNum has 1 row
22        If NR_TV > 1 And NR_RN = 1 Then
23            WriteRowIncrements = True
24            If RowNum(1, 1) + NR_TV - 1 > NR_TA Then Throw "Cannot write beyond the bottom row of TheArray"
25        End If
26        If NR_TV > 1 And NR_RN > 1 And NR_TV <> NR_RN Then
27            Throw "TheValue and RowNum are not conformable. If they both have more than one row then they must have the same number of rows"
28        End If

29        If NC_TV > 1 And NC_CN = 1 Then
30            WriteColIncrements = True
31            If ColNum(1, 1) + NC_TV - 1 > NC_TA Then Throw "Cannot write beyond the right column of TheArray"
32        End If
33        If NC_TV > 1 And NC_CN > 1 And NC_TV <> NC_CN Then
34            Throw "TheValue and ColNum are not conformable. If they both have more than one column then they must have the same number of columns"
35        End If

          Dim OpCode As Long
36        Select Case LCase(Operation)
              Case "overwrite"
37                OpCode = 0
38            Case "add"
39                OpCode = 1
40            Case "subtract"
41                OpCode = 2
42            Case "multiply"
43                OpCode = 3
44            Case "divide"
45                OpCode = 4
46            Case "concatenate"
47                OpCode = 5
48            Case Else
49                Throw "Operation not recognised must be 'Overwrite' (or omitted), 'Add', 'Subtract', 'Multiply', 'Divide' or 'Concatenate'"
50        End Select

          Dim ColLockForReading As Boolean
          Dim ColLockForWriting As Boolean
          Dim iLoopsTo As Long
          Dim jLoopsTo As Long
          Dim ReadCol As Long
          Dim ReadRow As Long
          Dim RowLockForReading As Boolean
          Dim RowLockForWriting As Boolean
          Dim WriteCol As Long
          Dim WriteRow As Long

51        ColLockForReading = NC_TV = 1
52        RowLockForReading = NR_TV = 1
53        ColLockForWriting = NC_CN = 1
54        RowLockForWriting = NR_RN = 1
55        iLoopsTo = Max(NR_TV, NR_RN)
56        jLoopsTo = Max(NC_TV, NC_CN)

57        For i = 1 To iLoopsTo
58            ReadRow = IIf(RowLockForReading, 1, i)
59            If WriteRowIncrements Then
60                WriteRow = RowNum(1, 1) + i - 1
61            Else
62                WriteRow = IIf(RowLockForWriting, RowNum(1, 1), RowNum(i, 1))
63            End If
64            For j = 1 To jLoopsTo
65                ReadCol = IIf(ColLockForReading, 1, j)
66                If WriteColIncrements Then
67                    WriteCol = ColNum(1, 1) + j - 1
68                Else
69                    WriteCol = IIf(ColLockForWriting, ColNum(1, 1), ColNum(1, j))
70                End If
71                Select Case OpCode
                      Case 0
72                        TheArray(WriteRow, WriteCol) = TheValue(ReadRow, ReadCol)
73                    Case 1
74                        TheArray(WriteRow, WriteCol) = SafeAdd(TheArray(WriteRow, WriteCol), TheValue(ReadRow, ReadCol))
75                    Case 2
76                        TheArray(WriteRow, WriteCol) = SafeSubtract(TheArray(WriteRow, WriteCol), TheValue(ReadRow, ReadCol))
77                    Case 3
78                        TheArray(WriteRow, WriteCol) = SafeMultiply(TheArray(WriteRow, WriteCol), TheValue(ReadRow, ReadCol))
79                    Case 4
80                        TheArray(WriteRow, WriteCol) = SafeDivide(TheArray(WriteRow, WriteCol), TheValue(ReadRow, ReadCol))
81                    Case 5
82                        TheArray(WriteRow, WriteCol) = CStr(TheArray(WriteRow, WriteCol)) & CStr(TheValue(ReadRow, ReadCol))
83                End Select
84            Next j
85        Next i

86        sInsert = TheArray

87        Exit Function
ErrHandler:
88        sInsert = "#sInsert (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function Max(x As Long, y As Long) As Long
1         If x > y Then
2             Max = x
3         Else
4             Max = y
5         End If
End Function
