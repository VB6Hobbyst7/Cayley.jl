Attribute VB_Name = "modArrayFnsF"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sPartialProduct
' Author    : Philip Swannell
' Date      : 16-Jun-2016
' Purpose   : Returns "partial product" or "running product" of an array of numbers. Each element in the
'             return is the product of the element in the corresponding position in the
'             input together with all elements directly above.
' Arguments
' TheArray  : An array of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sPartialProduct(ByVal TheArray As Variant)
Attribute sPartialProduct.VB_Description = "Returns ""partial product"" or ""running product"" of an array of numbers. Each element in the return is the product of the element in the corresponding position in the input together with all elements directly above."
Attribute sPartialProduct.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         ReDim Result(1 To N, 1 To M)
4         For j = 1 To M
5             Result(1, j) = TheArray(1, j)
6             For i = 2 To N
7                 Result(i, j) = SafeMultiply(Result(i - 1, j), TheArray(i, j))
8             Next i
9         Next j
10        sPartialProduct = Result
11        Exit Function
ErrHandler:
12        sPartialProduct = "#sPartialProduct (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sPartialSum
' Author    : Philip Swannell
' Date      : 30-Apr-2015
' Purpose   : Returns "partial sum" or "running total" of an array of numbers. Each element in the
'             return is the sum of the element in the corresponding position in the input
'             together with all elements directly above.
' Arguments
' TheArray  : An array of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sPartialSum(ByVal TheArray As Variant)
Attribute sPartialSum.VB_Description = "Returns ""partial sum"" or ""running total"" of an array of numbers. Each element in the return is the sum of the element in the corresponding position in the input together with all elements directly above."
Attribute sPartialSum.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         ReDim Result(1 To N, 1 To M)
4         For j = 1 To M
5             Result(1, j) = TheArray(1, j)
6             For i = 2 To N
7                 Result(i, j) = SafeAdd(Result(i - 1, j), TheArray(i, j))
8             Next i
9         Next j
10        sPartialSum = Result
11        Exit Function
ErrHandler:
12        sPartialSum = "#sPartialSum (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SpeedTestArrayTranspose
' Author    : Philip Swannell
' Date      : 21-May-2015
' Purpose   : Quick test harness for speed and accuracy. Speed seems similar, sArrayTranspose
'             copes better with very large arrays - WorksheetFunction.ArrayTranspose
'             fails silently (return is wrong size) for 1xN array for N > 65536
' -----------------------------------------------------------------------------------------------------------------------
Sub SpeedTestArrayTranspose()
          Dim N As Long
          Dim Res
          Dim Res2
          Dim res3
          Dim t1
          Dim t2
          Dim t3
1         N = 65536

2         Res = sReshape(sIntegers(17), N, 1)
3         t1 = sElapsedTime()
4         Res2 = Application.WorksheetFunction.Transpose(Res)
5         t2 = sElapsedTime()
6         res3 = sArrayTranspose(Res)
7         t3 = sElapsedTime()
8         Debug.Print "NumElements=" & Format$(sNRows(Res) * sNCols(Res), "###,###"), "WorksheetFunction", t2 - t1, "sArrayTranspose", t3 - t2, sArraysIdentical(Res2, res3)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegExpFromLiteral
' Author    : Philip Swannell
' Date      : 7-Dec-2015
' Purpose   : Returns a regular expression that matches a given string literal, by inserting backslash
'             characters when necessary to escape so-called metacharacters i.e. characters
'             in the list \.$^{}[]()|*+?
' Arguments
' StringLiteral: The literal string to match against. Example: if StringLiteral is \$ then the return from
'             the function will be\\\$. May also be an array of strings.
' Invert    : FALSE to convert a literal string into a regular expression that matches it. TRUE to
'             "undo" such conversion, recovering the literal string. Thus
'             sRegExpFromLiteral(sRegExpFromLiteral(x,FALSE),TRUE) = x for all strings x.
' -----------------------------------------------------------------------------------------------------------------------
Function sRegExpFromLiteral(ByVal StringLiteral As Variant, Optional Invert As Boolean)
Attribute sRegExpFromLiteral.VB_Description = "Returns a regular expression that matches a given string literal, by inserting backslash characters when necessary to escape so-called metacharacters i.e. characters in the list \\.$^{}[]()|*+?"
Attribute sRegExpFromLiteral.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim Res() As Variant

1         On Error GoTo ErrHandler
2         If TypeName(StringLiteral) = "Range" Then
3             StringLiteral = StringLiteral.Value2
4         End If

5         If VarType(StringLiteral) < vbArray Then
6             sRegExpFromLiteral = RegExpFromLiteral(CStr(StringLiteral), Invert)
7         ElseIf IsArray(StringLiteral) Then
8             Select Case NumDimensions(StringLiteral)
                  Case 1
9                     ReDim Res(LBound(StringLiteral) To UBound(StringLiteral))
10                    For i = LBound(StringLiteral) To UBound(StringLiteral)
11                        Res(i) = RegExpFromLiteral(CStr(StringLiteral(i)), Invert)
12                    Next i
13                Case 2
14                    ReDim Res(LBound(StringLiteral, 1) To UBound(StringLiteral, 1), LBound(StringLiteral, 2) To UBound(StringLiteral, 2))
15                    For i = LBound(StringLiteral, 1) To UBound(StringLiteral, 1)
16                        For j = LBound(StringLiteral, 2) To UBound(StringLiteral, 2)
17                            Res(i, j) = RegExpFromLiteral(CStr(StringLiteral(i, j)), Invert)
18                        Next j
19                    Next i
20                Case Else
21                    Throw "Arrays of dimension more than 2 are not handled"
22            End Select
23            sRegExpFromLiteral = Res
24        End If

25        Exit Function
ErrHandler:
26        sRegExpFromLiteral = "#sRegExpFromLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegExReplace
' Author    : Philip Swannell
' Date      : 13-Jan-2016
' Purpose   : Uses regular expressions to make replacement in a set of input strings.
'
'             The function replaces every instance of the regular expression match with the
'             replacement.
' Arguments
' InputString: Input string to be transformed. Can be an array. Non-string elements will be left
'             unchanged.
' RegularExpression: A standard regular expression string.
' Replacement: A replacement template for each match of the regular expression in the input string.
' CaseSensitive: Whether matching should be case-sensitive (TRUE) or not (FALSE).
'
' Notes     : Details of regular expressions are given under sIsRegMatch. The replacement string can be
'             an explicit string, and it can also contain special escape sequences that are
'             replaced by the characters they represent. The options available are:
'
'             Characters Replacement
'             $n        n-th backreference. That is, a copy of the n-th matched group
'             specified with parentheses in the regular expression. n must be an integer
'             value designating a valid backreference, greater than zero, and of two digits
'             at most.
'             $&       A copy of the entire match
'             $`        The prefix, that is, the part of the target sequence that precedes
'             the match.
'             $´        The suffix, that is, the part of the target sequence that follows
'             the match.
'             $$        A single $ character.
' -----------------------------------------------------------------------------------------------------------------------
Function sRegExReplace(InputString As Variant, RegularExpression As String, Replacement As String, Optional CaseSensitive As Boolean)
Attribute sRegExReplace.VB_Description = "Uses regular expressions to make replacement in a set of input strings.\n\nThe function replaces every instance of the regular expression match with the replacement. "
Attribute sRegExReplace.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim Result() As String
          Dim rx As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             sRegExReplace = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If

6         Set rx = New RegExp

7         With rx
8             .IgnoreCase = Not (CaseSensitive)
9             .Pattern = RegularExpression
10            .Global = True
11        End With

12        If VarType(InputString) = vbString Then
13            sRegExReplace = rx.Replace(InputString, Replacement)
14            GoTo Cleanup
15        ElseIf VarType(InputString) < vbArray Then
16            sRegExReplace = InputString
17            GoTo Cleanup
18        End If
19        If TypeName(InputString) = "Range" Then InputString = InputString.Value2

20        Select Case NumDimensions(InputString)
              Case 2
21                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1), LBound(InputString, 2) To UBound(InputString, 2))
22                For i = LBound(InputString, 1) To UBound(InputString, 1)
23                    For j = LBound(InputString, 2) To UBound(InputString, 2)
24                        If VarType(InputString(i, j)) = vbString Then
25                            Result(i, j) = rx.Replace(InputString(i, j), Replacement)
26                        Else
27                            Result(i, j) = InputString(i, j)
28                        End If
29                    Next j
30                Next i
31            Case 1
32                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1))
33                For i = LBound(InputString, 1) To UBound(InputString, 1)
34                    If VarType(InputString(i)) = vbString Then
35                        Result(i) = rx.Replace(InputString(i), Replacement)
36                    Else
37                        Result(i) = InputString(i)
38                    End If
39                Next i
40            Case Else
41                Throw "InputString must be a String or an array with 1 or 2 dimensions"
42        End Select
43        sRegExReplace = Result

Cleanup:
44        Set rx = Nothing
45        Exit Function
ErrHandler:
46        sRegExReplace = "#sRegExReplace (line " & CStr(Erl) + "): " & Err.Description & "!"
47        Set rx = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegMatch
' Author    : Philip Swannell
' Date      : 12-Jun-2017
' Purpose   : Implements Regular Expressions exposed by "Microsoft VBScript Regular Expressions 5.5".
'             The function returns the first sub-string of StringToSearch that matches
'             RegularExpression, or FALSE if there are zero matches.
' Arguments
' RegularExpression: The regular expression. Must be a string. Example cat|dog to match on either the string
'             cat or the string dog.
' StringToSearch: The string to match. May be an array in which case the return from the function is an
'             array of the same dimensions.
' CaseSensitive: TRUE for case-sensitive matching, FALSE for case-insensitive matching. This argument is
'             optional, defaulting to FALSE for case-insensitive matching.
' -----------------------------------------------------------------------------------------------------------------------
Function sRegMatch(RegularExpression As String, ByVal StringToSearch As Variant, Optional CaseSensitive As Boolean = False)
Attribute sRegMatch.VB_Description = "Implements Regular Expressions exposed by ""Microsoft VBScript Regular Expressions 5.5"". The function returns the first sub-string of StringToSearch that matches RegularExpression, or FALSE if there are zero matches."
Attribute sRegMatch.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim MatchRes As IMatchCollection2
          Dim Result() As Variant
          Dim rx As VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             sRegMatch = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If

6         Set rx = New RegExp
7         With rx
8             .IgnoreCase = Not CaseSensitive
9             .Pattern = RegularExpression
10            .Global = False        'Find first match only
11        End With

12        If VarType(StringToSearch) = vbString Then
13            Set MatchRes = rx.Execute(StringToSearch)
14            If MatchRes.Count > 0 Then
15                sRegMatch = MatchRes(0).Value
16            Else
17                sRegMatch = "#No match found!"
18            End If
19            GoTo EarlyExit
20        ElseIf VarType(StringToSearch) < vbArray Then
21            sRegMatch = "#StringToSearch must be a string!"
22            GoTo EarlyExit
23        End If
24        If TypeName(StringToSearch) = "Range" Then StringToSearch = StringToSearch.Value2

25        Select Case NumDimensions(StringToSearch)
              Case 2
26                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1), LBound(StringToSearch, 2) To UBound(StringToSearch, 2))
27                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
28                    For j = LBound(StringToSearch, 2) To UBound(StringToSearch, 2)
29                        If VarType(StringToSearch(i, j)) = vbString Then
30                            Set MatchRes = rx.Execute(StringToSearch(i, j))
31                            If MatchRes.Count > 0 Then
32                                Result(i, j) = MatchRes(0).Value
33                            Else
34                                sRegMatch = "#No match found!"
35                            End If
36                        Else
37                            Result(i, j) = "#StringToSearch must be a string!"
38                        End If
39                    Next j
40                Next i
41            Case 1
42                ReDim Result(LBound(StringToSearch, 1) To UBound(StringToSearch, 1))
43                For i = LBound(StringToSearch, 1) To UBound(StringToSearch, 1)
44                    If VarType(StringToSearch(i)) = vbString Then
45                        Set MatchRes = rx.Execute(StringToSearch(i))
46                        If MatchRes.Count > 0 Then
47                            Result(i) = MatchRes(0).Value
48                        Else
49                            sRegMatch = "#No match found!"
50                        End If
51                    Else
52                        Result(i) = "#StringToSearch must be a string!"
53                    End If
54                Next i
55            Case Else
56                Throw "StringToSearch must be String or array with 1 or 2 dimensions"
57        End Select

58        sRegMatch = Result
EarlyExit:
59        Set rx = Nothing

60        Exit Function
ErrHandler:
61        sRegMatch = "#sRegMatch (line " & CStr(Erl) + "): " & Err.Description & "!"
62        Set rx = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRemoveDuplicateRows
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : Returns an array listing the unique rows in the input array. Optionally can sort the
'             return as well, in which case the return is sorted in ascending order keying
'             on all columns from left to right.
' Arguments
' InputArray: The array from which to remove duplicated rows.
' SortAsWell: TRUE if the array is to be returned in sorted order. If omitted defaults to FALSE. Note
'             that the function executes faster if this argument is passed as TRUE.
' CaseSensitive: TRUE if comparison of strings to determine duplication is to be case-sensitive. If omitted
'             defaults to FALSE for case insensitive comparison of strings.
' -----------------------------------------------------------------------------------------------------------------------
Function sRemoveDuplicateRows(ByVal InputArray As Variant, _
        Optional SortAsWell As Boolean, _
        Optional CaseSensitive As Boolean = False)
Attribute sRemoveDuplicateRows.VB_Description = "Returns an array listing the unique rows in the input array. Optionally can sort the return as well, in which case the return is sorted in ascending order keying on all columns from left to right. "
Attribute sRemoveDuplicateRows.VB_ProcData.VB_Invoke_Func = " \n27"

          Dim ChooseVector As Variant
          Dim fromCol As Long
          Dim i As Long
          Dim j As Long
          Dim N As Long
          Dim TempArray() As Variant
          Dim ToCol As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR InputArray

3         N = sNRows(InputArray)

4         If Not SortAsWell Then
5             fromCol = 2
6             TempArray = sSortedArrayByAllCols(sArrayRange(sIntegers(N), InputArray), True, CaseSensitive)
7         Else
8             fromCol = 1
9             TempArray = sSortedArrayByAllCols(InputArray, False, CaseSensitive)
10        End If

11        ToCol = sNCols(TempArray)

12        ChooseVector = sReshape(False, N, 1)
13        ChooseVector(1, 1) = True
14        For i = 2 To N
15            For j = fromCol To ToCol
16                If Not sEquals(TempArray(i, j), TempArray(i - 1, j), CaseSensitive) Then
17                    ChooseVector(i, 1) = True
18                    Exit For
19                End If
20            Next j
21        Next i
22        TempArray = sMChoose(TempArray, ChooseVector)

23        If Not SortAsWell Then
              'Have to put back into original order
24            TempArray = sSortedArray(TempArray, 1, , , True)
25            TempArray = sSubArray(TempArray, 1, 2)
26        End If

27        sRemoveDuplicateRows = TempArray

28        Exit Function
ErrHandler:
29        Throw "#sRemoveDuplicateRows (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowMedian
' Author    : Philip Swannell
' Date      : 15-Jul-2017
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             return is the median of the corresponding row of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' IgnoreErrors: If FALSE (the default) then error values (#N/A!, #REF! etc) in ArrayOfNumbers yield error
'             strings in the corresponding element of the return. If TRUE error values are
'             excluded from calculation of the column medians. String values are always
'             excluded.
' -----------------------------------------------------------------------------------------------------------------------
Function sRowMedian(ByVal ArrayOfNumbers, Optional IgnoreErrors As Boolean)
Attribute sRowMedian.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the return is the median of the corresponding row of the input."
Attribute sRowMedian.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As Variant
2         Force2DArrayR ArrayOfNumbers, NR, NC
          Dim tmp() As Variant

3         ReDim Result(1 To NR, 1 To 1)
4         ReDim tmp(1 To NC)
5         If NC = 1 Then
6             Result = ArrayOfNumbers
7             For i = 1 To NR
8                 If Not IsNumberOrDate(Result(i, 1)) Then
9                     Result(i, 1) = "#Cannot calculate median!"
10                End If
11            Next i
12        Else
13            For i = 1 To NR
14                If IgnoreErrors Then
15                    For j = 1 To NC
16                        If IsNumberOrDate(ArrayOfNumbers(i, j)) Then
17                            tmp(j) = ArrayOfNumbers(i, j)
18                        Else
19                            tmp(j) = vbNullString
20                        End If
21                    Next j
22                Else
23                    For j = 1 To NC
24                        tmp(j) = ArrayOfNumbers(i, j)
25                    Next j
26                End If
27                Result(i, 1) = SafeMedian(tmp)
28            Next i
29        End If
30        sRowMedian = Result
31        Exit Function
ErrHandler:
32        sRowMedian = "#sRowMedian (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sStringsMatchingRegExp
' Author    : Philip Swannell
' Date      : 14-Apr-2018, 23-Apr-2021
' Purpose   : Returns the elements of Strings that match RegularExpression. Elements of Strings that are
'             not strings (i.e are numbers, logicals, etc.) are excluded from the return.
' Arguments
' Strings   : An array of strings. The return has a single column even if Strings has more than one
'             column.
' RegularExpression: A valid regular expression. See sIsRegMatch.
' CaseSensitive: Whether regular expression match is case sensitive or not.
' -----------------------------------------------------------------------------------------------------------------------
Function sStringsMatchingRegExp(ByVal Strings, RegularExpression As String, Optional CaseSensitive As Boolean = False)
Attribute sStringsMatchingRegExp.VB_Description = "Returns the elements of Strings that match RegularExpression. Elements of Strings that are not strings (i.e are numbers, logicals, etc.) are excluded from the return."
Attribute sStringsMatchingRegExp.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim ChooseVector
1         On Error GoTo ErrHandler
2         Force2DArrayR Strings
3         If sNCols(Strings) > 1 Then
4             Strings = sReshape(Strings, sNRows(Strings) * sNCols(Strings), 1)
5         End If
6         ChooseVector = sArrayIf(sArrayIsText(Strings), ThrowIfError(sIsRegMatch(RegularExpression, Strings, CaseSensitive)), False)
7         sStringsMatchingRegExp = sMChoose(Strings, ChooseVector)
8         Exit Function
ErrHandler:
9         sStringsMatchingRegExp = "#sStringsMatchingRegExp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSumOfNums
' Author    : Philip Swannell
' Date      : 28-May-2015
' Purpose   : Adds up all the numerical elements of the input TheArray, ignoring those elements which
'             are not numeric. If no elements are numeric, the number zero is returned.
'             Note that strings containing numbers, such as "123", will be ignored.
' Arguments
' TheArray  : Input array of elements of any type. Only numeric elements will be used in the sum.
' -----------------------------------------------------------------------------------------------------------------------
Function sSumOfNums(ByVal TheArray)
Attribute sSumOfNums.VB_Description = "Adds up all elements of the TheArray, ignoring those which are not numbers. If no elements are numbers, zero is returned. Strings that look like numbers, such as ""123"", are ignored."
Attribute sSumOfNums.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim c As Variant
          Dim Result As Double

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray

3         For Each c In TheArray
4             If IsNumberOrDate(c) Then
5                 Result = Result + c
6             End If
7         Next c

8         sSumOfNums = Result
9         Exit Function
ErrHandler:
10        sSumOfNums = "#sSumOfNums (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function sAverageOfNums(ByVal TheArray)
Attribute sAverageOfNums.VB_Description = "Returns the average of the elements of TheArray, ignoring elements that are not numbers. If no elements are numbers, an error string is returned."
Attribute sAverageOfNums.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim c As Variant
          Dim N As Long
          Dim Result As Double

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray

3         For Each c In TheArray
4             If IsNumberOrDate(c) Then
5                 Result = Result + c
6                 N = N + 1
7             End If
8         Next c
9         If N = 0 Then
10            sAverageOfNums = "#No numbers found!"
11        Else
12            sAverageOfNums = Result / N
13        End If
14        Exit Function
ErrHandler:
15        sAverageOfNums = "#sAverageOfNums (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSuppressNAs
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : Replaces each instance of #N/A in TheArray with a null string. If sSuppressNAs(TheArray)
'             is entered as an array formula in a larger range of cells than TheArray the
'             additional rows/columns are populated with null strings.
' Arguments
' TheArray  : An array of arbitrary values.
' ReplaceWith: The value to replace each #N/A. This argument is optional and defaults to the null string.
' -----------------------------------------------------------------------------------------------------------------------
Function sSuppressNAs(ByVal TheArray, Optional ReplaceWith = vbNullString)
Attribute sSuppressNAs.VB_Description = "Replaces each instance of #N/A! in TheArray with a null string. If sSuppressNAs(TheArray) is entered as an array formula in a larger range of cells than TheArray the additional rows/columns are populated with null strings."
Attribute sSuppressNAs.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NCOut As Long
          Dim NR As Long
          Dim NROut As Long
          Dim RangeNC As Long
          Dim RangeNR As Long
          Dim Result()

1         On Error GoTo ErrHandler
2         Select Case VarType(ReplaceWith)
              Case vbString, vbBoolean, vbDouble, vbInteger, vbSingle, vbLong, vbDate, vbError
                  'OK
3             Case Else
4                 Throw "ReplaceWith must be a string or number or logical value"
5         End Select

6         If TypeName(Application.Caller) = "Range" Then
7             RangeNR = Application.Caller.Rows.Count
8             RangeNC = Application.Caller.Columns.Count
9         End If
10        Force2DArrayR TheArray, NR, NC
11        NROut = SafeMax(NR, RangeNR)
12        NCOut = SafeMax(NC, RangeNC)
13        ReDim Result(1 To NROut, 1 To NCOut)

14        For i = 1 To NR
15            For j = 1 To NC
16                Result(i, j) = CoreSuppressNAs(TheArray(i, j), ReplaceWith)
17            Next j
18            For j = NC + 1 To NCOut
19                Result(i, j) = ReplaceWith
20            Next j
21        Next i
22        For i = NR + 1 To NROut
23            For j = 1 To NCOut
24                Result(i, j) = ReplaceWith
25            Next j
26        Next i
27        sSuppressNAs = Result

28        Exit Function
ErrHandler:
29        sSuppressNAs = "#sSuppressNAs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sTokeniseString
' Author    : Philip Swannell
' Date      : 07-Oct-2013
' Purpose   : Breaks up TheString into sub-strings with breaks at the positions at which the Delimiter
'             character appears, and returns the sub-strings as a column array.
'             Example: sTokeniseString("Goodbye cruel world!", " ") returns the array
'             "Goodbye"
'             "cruel"
'             "world!"
' Arguments
' TheString : The string to be tokenised.
' Delimiter : The delimiter character, can be multiple characters. The search for the delimiter
'             character is case insensitive.
' -----------------------------------------------------------------------------------------------------------------------
Function sTokeniseString(TheString As String, Optional Delimiter As String = ",")
Attribute sTokeniseString.VB_Description = "Breaks up TheString into sub-strings with breaks at the positions at which the Delimiter character appears, and returns the sub-strings as a column array."
Attribute sTokeniseString.VB_ProcData.VB_Invoke_Func = " \n25"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim LB As Long
          Dim N As Long
          Dim OneDArray
          Dim Res()
          Dim UB As Long
          'Changed PGS 10 July 2017 to cope when TheString had 91 million characters with 2338 delimiters, was previously using Application.WorksheetFunction.Transpose, but that failed in this case
2         OneDArray = VBA.Split(TheString, Delimiter, -1, vbTextCompare)
3         LB = LBound(OneDArray): UB = UBound(OneDArray)
4         N = UB - LB + 1
5         ReDim Res(1 To N, 1 To 1)
6         For i = 1 To N
7             Res(i, 1) = OneDArray(i - 1)
8         Next
9         sTokeniseString = Res
10        Exit Function
ErrHandler:
11        sTokeniseString = "#sTokeniseString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sTokeniseStrings
' Author    : Philip Swannell
' Date      : 14-Apr-2018
' Purpose   : Parses a vector of delimited strings into an array.
' Arguments
' TheStrings: The strings to be tokenised, must have either one column or one row. If not all of the
'             strings have the same number of tokens then the result is padded with #N/A
'             values. No attempt is made to tokenise non-strings, which appear unchanged in
'             the output.
' Delimiter : The delimiter character, can be multiple characters and if omited defaults to a comma. The
'             search for the delimiter character is case sensitive.
'
' Notes     : Example:
'             If the range A1:A3 contains:
'             "a,b,c"
'             "d,e,f"
'             "g,h,i"
'             then =sTokeniseStrings(A1:A3,",") would return a 3x3 array:
'             "a"   "b"   "c"
'             "d"   "e"   "f"
'             "g"   "h"   "i"
' -----------------------------------------------------------------------------------------------------------------------
Function sTokeniseStrings(ByVal TheStrings As Variant, Optional Delimiter As String = ",")
Attribute sTokeniseStrings.VB_Description = "Parses a vector of delimited strings into an array."
Attribute sTokeniseStrings.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim IsCol As Boolean
          Dim j
          Dim LB
          Dim MaxElements As Long
          Dim N As Long
          Dim Result
          Dim tmp
          Dim x As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR TheStrings
3         If sNRows(TheStrings) = 1 Then
4         ElseIf sNCols(TheStrings) = 1 Then
5             IsCol = True
6         Else
7             Throw "TheStrings must have either one row or one column"
8         End If
9         MaxElements = 1
10        For Each x In TheStrings
11            If VarType(x) = vbString Then
12                N = Len(CStr(x)) - Len(Replace(CStr(x), Delimiter, vbNullString))
13                If N + 1 > MaxElements Then
14                    MaxElements = N + 1
15                End If
16            End If
17        Next

18        If IsCol Then
19            Result = sReshape(CVErr(xlErrNA), sNRows(TheStrings), MaxElements)
20            For i = 1 To sNRows(TheStrings)
21                If VarType(TheStrings(i, 1)) = vbString Then
22                    tmp = VBA.Split(CStr(TheStrings(i, 1)), Delimiter)
23                    LB = LBound(tmp)
24                    For j = LB To UBound(tmp)
25                        Result(i, 1 - LB + j) = tmp(j)
26                    Next j
27                Else
28                    Result(i, 1) = TheStrings(i, 1)
29                End If
30            Next i
31        Else
32            Result = sReshape(CVErr(xlErrNA), MaxElements, sNCols(TheStrings))
33            For i = 1 To sNCols(TheStrings)
34                If VarType(TheStrings(1, i)) = vbString Then
35                    tmp = VBA.Split(CStr(TheStrings(1, i)), Delimiter)
36                    LB = LBound(tmp)
37                    For j = LB To UBound(tmp)
38                        Result(1 - LB + j, i) = tmp(j)
39                    Next j
40                Else
41                    Result(1, i) = TheStrings(1, i)
42                End If
43            Next i
44        End If
45        sTokeniseStrings = Result
46        Exit Function
ErrHandler:
47        sTokeniseStrings = "#sTokeniseStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sVlookup
' Author    : Philip Swannell
' Date      : 09-Apr-2020
' Purpose   : Searches a column of Table for LookupValues and returns the corresponding elements from
'             another column of Table. Like VLOOKUP except: matching is always exact;
'             columns may be specified as header strings; the column searched need not be
'             the first column.
' Arguments
' LookupValues: The value to search for. May be an array of values.
' Table     : An array of arbitrary values.
' ColNoOrHeader: Either a) the number of the column of Table from which values are to be returned; or b) a
'             "Header string" matching an element in the top row of Table; or c) 1-row
'             array of column numbers or header strings. If omitted defaults to 2 for 2nd
'             column of Table.
' KeyColNoOrHeader: Either a) the number of the column of Table to be searched for LookupValues; or b) a
'             "Header string" matching an element in the top row of Table. If omitted,
'             defaults to 1, for the leftmost column of table.
' -----------------------------------------------------------------------------------------------------------------------
Function sVLookup(ByVal LookupValues, ByVal Table, Optional ByVal ColNoOrHeader = 2, Optional KeyColNoOrHeader = 1)
Attribute sVLookup.VB_Description = "Searches a column of Table for LookupValues and returns the corresponding elements from another column of Table. Like VLOOKUP except: matching is always exact; columns may be specified as header strings; the column searched need not be the first column."
Attribute sVLookup.VB_ProcData.VB_Invoke_Func = " \n27"
1         On Error GoTo ErrHandler
2         If VarType(ColNoOrHeader) < vbArray Then
3             sVLookup = sVLookupCore(LookupValues, Table, ColNoOrHeader, KeyColNoOrHeader)
4         Else
5             If NumDimensions(ColNoOrHeader) = 1 Then
6                 Force2DArrayR ColNoOrHeader
7             ElseIf sNRows(ColNoOrHeader) > 1 Then
8                 Throw "ColNoOrHeader must be a single value or a one-row array of values"
9             End If
              Dim Result
              Dim ThisResult
              Dim i As Long, j As Long, k As Long, WriteCol As Long
10            WriteCol = 0
11            ThisResult = sVLookupCore(LookupValues, Table, ColNoOrHeader(1, 1), KeyColNoOrHeader)
12            If VarType(ThisResult) < vbArray Then ThrowIfError (ThisResult)
13            WriteCol = WriteCol + sNCols(ThisResult)
14            If sNCols(ColNoOrHeader) = 1 Then
15                sVLookup = ThisResult
16                Exit Function
17            Else
18                Result = sArrayRange(ThisResult, sReshape("", sNRows(ThisResult), sNCols(ThisResult) * (sNCols(ColNoOrHeader) - 1)))
19                For k = 2 To sNCols(ColNoOrHeader)
20                    ThisResult = sVLookupCore(LookupValues, Table, ColNoOrHeader(1, k), KeyColNoOrHeader)
21                    If VarType(ThisResult) < vbArray Then
22                        ThrowIfError ThisResult
23                    End If
24                    Force2DArray ThisResult
                      
25                    For i = 1 To sNRows(ThisResult)
26                        For j = 1 To sNCols(ThisResult)
27                            Result(i, j + WriteCol) = ThisResult(i, j)
28                        Next
29                    Next
30                    WriteCol = WriteCol + sNCols(ThisResult)
31                Next
32            End If
33            sVLookup = Result
34        End If
35        Exit Function
ErrHandler:
36        sVLookup = "#sVLookup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function sVLookupCore(ByVal LookupValues, ByVal Table, Optional ColNoOrHeader = 2, Optional KeyColNoOrHeader = 1)
Attribute sVLookupCore.VB_Description = "Searches a column of Table for LookupValues and returns the corresponding elements from another column of Table. Like VLOOKUP except: matching is always exact; columns may be specified as header strings; the column searched need not be the first column."
Attribute sVLookupCore.VB_ProcData.VB_Invoke_Func = " \n25"

          Dim HeaderRowT As Variant
          Dim i As Long
          Dim KeyColNo As Long
          Dim KeyColumn
          Dim LookupColNo As Long
          Dim MatchRes
          Dim NC As Long
          Dim Res As Variant
          Dim Result() As Variant
          Dim ValuesColumn
          Const ErrorLookupColNoOrHeader = "ColNoOrHeader must be given as a number or as a string matching a string in the first row of Table"
          Const ErrorKeyColNoOrHeader = "KeyColNoOrHeader must be given as a number or as a string matching a string in the first row of Table"

1         On Error GoTo ErrHandler

2         If Not (IsNumber(ColNoOrHeader)) And (VarType(ColNoOrHeader) <> vbString) Then
3             Throw ErrorLookupColNoOrHeader
4         End If

5         If Not (IsNumber(KeyColNoOrHeader)) And (VarType(KeyColNoOrHeader) <> vbString) Then
6             Throw ErrorKeyColNoOrHeader
7         End If

8         If VarType(ColNoOrHeader) = vbString Or VarType(KeyColNoOrHeader) = vbString Then
9             HeaderRowT = sArrayTranspose(sSubArray(Table, 1, 1, 1))
10        End If
11        If VarType(ColNoOrHeader) = vbString Then
12            Res = sMatch(ColNoOrHeader, HeaderRowT)
13            If Not IsNumber(Res) Then
14                Throw "Cannot find " + ColNoOrHeader + " in top row of Table"
15            Else
16                LookupColNo = Res
17            End If
18        Else
19            LookupColNo = ColNoOrHeader
20        End If
21        If VarType(KeyColNoOrHeader) = vbString Then
22            Res = sMatch(KeyColNoOrHeader, HeaderRowT)
23            If Not IsNumber(Res) Then
24                Throw "Cannot find " + KeyColNoOrHeader + " in top row of Table"
25            Else
26                KeyColNo = Res
27            End If
28        Else
29            KeyColNo = KeyColNoOrHeader
30        End If

31        NC = sNCols(Table)

32        If LookupColNo < 1 Or LookupColNo > NC Then Throw ErrorLookupColNoOrHeader
33        If KeyColNo < 1 Or KeyColNo > NC Then Throw ErrorKeyColNoOrHeader

34        If TypeName(Table) = "Range" And VarType(LookupValues) < vbArray Then        'Common use-case - looking up a singleton in a range, this branch of the code approx 2 times faster
35            On Error Resume Next
36            Res = Application.WorksheetFunction.Match(LookupValues, Table.Columns(KeyColNo), False)
37            On Error GoTo ErrHandler
38            If IsEmpty(Res) Then
39                sVLookupCore = "#Not found!"
40                Exit Function
41            Else
42                sVLookupCore = Table.Cells(Res, LookupColNo).Value
43            End If
44        Else
45            Force2DArrayRMulti LookupValues, Table

              Dim DoReshape As Boolean
              Dim LUVNC As Long
              Dim LUVNR As Long
46            LUVNR = sNRows(LookupValues): LUVNC = sNCols(LookupValues)
47            If LUVNC > 1 Then
48                DoReshape = True
49                LookupValues = sReshape(LookupValues, LUVNR * LUVNC, 1)
50            End If

51            ValuesColumn = sSubArray(Table, 1, LookupColNo, , 1)
52            KeyColumn = sSubArray(Table, 1, KeyColNo, , 1)
53            MatchRes = sMatch(LookupValues, KeyColumn)
54            If VarType(MatchRes) = vbString Then
55                sVLookupCore = "#Not found!"
56                Exit Function
57            ElseIf IsNumber(MatchRes) Then
58                sVLookupCore = ValuesColumn(MatchRes, 1)
59                Exit Function
60            ElseIf VarType(MatchRes) >= vbArray Then
                  Dim ResultNR As Long
61                ResultNR = sNRows(MatchRes)
62                ReDim Result(1 To ResultNR, 1 To 1)
63                For i = 1 To ResultNR
64                    If IsNumber(MatchRes(i, 1)) Then
65                        Result(i, 1) = ValuesColumn(MatchRes(i, 1), 1)
66                    Else
67                        Result(i, 1) = "#Not found!"
68                    End If
69                Next i
70                If DoReshape Then
71                    Result = sReshape(Result, LUVNR, LUVNC)
72                End If
73                sVLookupCore = Result
74            End If
75        End If

76        Exit Function
ErrHandler:
77        sVLookupCore = "#sVLookupCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub sVLookupTest()
          Dim LookupValues
          Dim Res1
          Dim Res2
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim Table

1         Table = sRandomVariable(100000, 2)
2         LookupValues = Table(100000, 1)
3         t1 = sElapsedTime()
4         Res1 = sVLookup(LookupValues, Table)
5         t2 = sElapsedTime()
6         Res2 = Application.WorksheetFunction.VLookup(LookupValues, Table, 2, False)
7         t3 = sElapsedTime()
8         Debug.Print "res1 = res2", Res1 = Res2, , "sVLookup", t2 - t1, "VLOOKUP", t3 - t2, "Ratio", (t2 - t1) / (t3 - t2)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sWordCount
' Author    : Philip Swannell
' Date      : 24-Sep-2018
' Purpose   : Counts the number of words in InputString. A word is defined as one or more consecutive
'             characters which are either alphanumeric (0-9, a-z, A-Z) or an underscore or
'             an accented character (ascii 138,140,142, 158, 159, 192 to 255)
' Arguments
' InputString: A string, or array of strings, in which case the return is an array of the same size.
' -----------------------------------------------------------------------------------------------------------------------
Function sWordCount(InputString As Variant)
Attribute sWordCount.VB_Description = "Counts the number of words in InputString. A word is defined as one or more consecutive characters which are either alphanumeric (0-9, a-z, A-Z) or an underscore or an accented character (ascii 138,140,142, 158, 159, 192 to 255)"
Attribute sWordCount.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim Result() As Long
          Dim rx1 As VBScript_RegExp_55.RegExp
          Dim rx2 As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler

2         Set rx1 = New RegExp
3         With rx1
4             .IgnoreCase = True
5             .Pattern = "[" & Chr$(138) & Chr$(140) & Chr$(142) & Chr$(158) & Chr$(159) & Chr$(192) & "-" & Chr$(255) & "]"
6             .Global = True
7         End With

8         Set rx2 = New RegExp
9         With rx2
10            .IgnoreCase = True
11            .Pattern = "\b"
12            .Global = True
13        End With

14        If VarType(InputString) = vbString Then
15            If Len(InputString) = 0 Then
16                sWordCount = 0
17            Else
18                sWordCount = (Len(rx2.Replace(rx1.Replace(InputString, "x"), "x")) - Len(InputString)) / 2
19            End If
20            GoTo Cleanup
21        ElseIf VarType(InputString) < vbArray Then
22            sWordCount = 0
23            GoTo Cleanup
24        End If
25        If TypeName(InputString) = "Range" Then InputString = InputString.Value2

26        Select Case NumDimensions(InputString)
              Case 2
27                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1), LBound(InputString, 2) To UBound(InputString, 2))
28                For i = LBound(InputString, 1) To UBound(InputString, 1)
29                    For j = LBound(InputString, 2) To UBound(InputString, 2)
30                        If VarType(InputString(i, j)) = vbString Then
31                            If Len(InputString(i, j)) = 0 Then
32                                Result(i, j) = 0
33                            Else
34                                Result(i, j) = (Len(rx2.Replace(InputString(i, j), "x")) - Len(InputString(i, j))) / 2
35                            End If
36                        Else
37                            Result(i, j) = 0
38                        End If
39                    Next j
40                Next i
41            Case 1
42                ReDim Result(LBound(InputString, 1) To UBound(InputString, 1))
43                For i = LBound(InputString, 1) To UBound(InputString, 1)
44                    If VarType(InputString(i)) = vbString Then
45                        If Len(InputString(i)) = 0 Then
46                            Result(i) = 0
47                        Else
48                            Result(i) = (Len(rx2.Replace(InputString(i), "x")) - Len(InputString(i))) / 2
49                        End If
50                    Else
51                        Result(i) = 0
52                    End If
53                Next i
54            Case Else
55                Throw "InputString must have no more than two dimensions"
56        End Select
57        sWordCount = Result
Cleanup:
58        Set rx2 = Nothing
59        Exit Function
ErrHandler:
60        sWordCount = "#sWordCount (line " & CStr(Erl) + "): " & Err.Description & "!"
61        Set rx2 = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSpillDown
' Author    : Philip Swannell
' Date      : 13-Nov-2020
' Purpose   : For use with dynamic array formulas. Returns a part of a spill range, from InputRange down
'             to the bottom of the spill range.
' Arguments
' InputRange: The top of the returned range, must be a cell or cells inside the spill range of a dynamic
'             array formula.
' -----------------------------------------------------------------------------------------------------------------------
Function sSpillDown(InputRange As Range)
Attribute sSpillDown.VB_Description = "For use with dynamic array formulas. Returns a part of a spill range, from InputRange down to the bottom of the spill range."
Attribute sSpillDown.VB_ProcData.VB_Invoke_Func = " \n31"

1         On Error GoTo ErrHandler
          Dim SRTopLeft As Range
          Dim SRTopLeft2 As Range
          Dim SR As Range
          
2         On Error Resume Next
3         Set SR = InputRange(1, 1).SpillParent.SpillingToRange
4         Set SRTopLeft = SR.Cells(1, 1)
5         If InputRange.Cells.Count > 1 Then
6             With InputRange
7                 Set SRTopLeft2 = .Cells(.Rows.Count, .Columns.Count).SpillParent
8             End With
9         Else
10            Set SRTopLeft2 = SRTopLeft
11        End If
12        On Error GoTo ErrHandler

13        If SRTopLeft Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
14        If SRTopLeft2 Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
15        If SRTopLeft.address <> SRTopLeft2.address Then Throw "InputRange is not a sub range of a dynamic array"
16        Set sSpillDown = InputRange.Resize(SR.row + SR.Rows.Count - InputRange.row)
17        Exit Function
ErrHandler:
18        sSpillDown = "#sSpillDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSpillRight
' Author    : Philip Swannell
' Date      : 13-Nov-2020
' Purpose   : For use with dynamic array formulas. Returns a part of a spill range, from InputRange
'             across to the right of the spill range.
' Arguments
' InputRange: The left of the returned range, must be a cell or cells inside the spill range of a
'             dynamic array formula.
' -----------------------------------------------------------------------------------------------------------------------
Function sSpillRight(InputRange As Range)
Attribute sSpillRight.VB_Description = "For use with dynamic array formulas. Returns a part of a spill range, from InputRange across to the right of the spill range."
Attribute sSpillRight.VB_ProcData.VB_Invoke_Func = " \n31"

1         On Error GoTo ErrHandler
          Dim SRTopLeft As Range
          Dim SRTopLeft2 As Range
          Dim SR As Range

2         On Error Resume Next
3         Set SR = InputRange(1, 1).SpillParent.SpillingToRange
4         Set SRTopLeft = SR.Cells(1, 1)
5         If InputRange.Cells.Count > 1 Then
6             With InputRange
7                 Set SRTopLeft2 = .Cells(.Rows.Count, .Columns.Count).SpillParent
8             End With
9         Else
10            Set SRTopLeft2 = SRTopLeft
11        End If
12        On Error GoTo ErrHandler

13        If SRTopLeft Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
14        If SRTopLeft2 Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
15        If SRTopLeft.address <> SRTopLeft2.address Then Throw "InputRange is not a sub range of a dynamic array"
            
16        Set sSpillRight = InputRange.Resize(, SR.Column + SR.Columns.Count - InputRange.Column)
17        Exit Function
ErrHandler:
18        sSpillRight = "#sSpillRight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSpillRightDown
' Author    : Philip Swannell
' Date      : 13-Nov-2020
' Purpose   : For use with dynamic array formulas. Returns a part of a spill range, from InputRange down
'             and across to the bottom right of the spill range.
' Arguments
' InputRange: The top left of the returned range, must be a cell inside the spill range of a dynamic
'             array formula.
' -----------------------------------------------------------------------------------------------------------------------
Function sSpillRightDown(InputRange As Range)
Attribute sSpillRightDown.VB_Description = "For use with dynamic array formulas. Returns a part of a spill range, from InputRange down and across to the bottom right of the spill range."
Attribute sSpillRightDown.VB_ProcData.VB_Invoke_Func = " \n31"

1         On Error GoTo ErrHandler
          Dim SRTopLeft As Range
          Dim SRTopLeft2 As Range
          Dim SR As Range
          
2         On Error Resume Next
3         Set SR = InputRange(1, 1).SpillParent.SpillingToRange
4         Set SRTopLeft = SR.Cells(1, 1)
5         If InputRange.Cells.Count > 1 Then
6             With InputRange
7                 Set SRTopLeft2 = .Cells(.Rows.Count, .Columns.Count).SpillParent
8             End With
9         Else
10            Set SRTopLeft2 = SRTopLeft
11        End If
12        On Error GoTo ErrHandler

13        If SRTopLeft Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
14        If SRTopLeft2 Is Nothing Then Throw "InputRange is not a sub range of a dynamic array"
15        If SRTopLeft.address <> SRTopLeft2.address Then Throw "InputRange is not a sub range of a dynamic array"
            
16        Set sSpillRightDown = InputRange.Resize(SR.row + SR.Rows.Count - InputRange.row, _
              SR.Column + SR.Columns.Count - InputRange.Column)
17        Exit Function
ErrHandler:
18        sSpillRightDown = "#sSpillRightDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sTrimNAs
' Author    : Philip Swannell
' Date      : 23-Apr-2021
' Purpose   : Removes all rows from the bottom of Data and columns at the right of Data that contain
'             only Excel error values or are empty.
' Arguments
' Data      : An array of data.
'
' Notes     : Example:
'             If the range A1:C3 contains:
'             1      2       #N/A
'             3      #N/A    #N/A
'             #N/A   #N/A    #N/A
'             then =sTrimNAs(A1:C3) would return a 2x2 array:
'             1      2
'             3      #N/A
' -----------------------------------------------------------------------------------------------------------------------
Function sTrimNAs(ByVal Data)
Attribute sTrimNAs.VB_Description = "Removes all rows from the bottom of Data and columns at the right of Data that contain only Excel error values or are empty."
Attribute sTrimNAs.VB_ProcData.VB_Invoke_Func = " \n27"

          Dim NR As Long, NC As Long, i As Long, j As Long, k As Long
          Dim LastRow As Long, LastCol As Long
            
1         LastRow = 1
2         LastCol = 1 ' because we cant return zero-sized array in VBA

3         On Error GoTo ErrHandler
4         Force2DArrayR Data, NR, NC

          'Diagonal scan from the bottom right
5         For k = 1 To NR * NC
6             If k = 1 Then
7                 i = NR: j = NC
8             Else
9                 If j < NC And i > 1 Then
10                    i = i - 1
11                    j = j + 1
12                ElseIf j = NC Then
13                    j = NC - (NR - i) - 1
14                    i = NR
15                    If j < 1 Then
16                        i = NR - (1 - j)
17                        j = 1
18                    End If
19                ElseIf i = 1 Then
20                    i = j - 1
21                    j = 1
22                    If i > NR Then
23                        j = 1 + i - NR
24                        i = NR
25                    End If
26                End If
27            End If
28            If Not (IsError(Data(i, j)) Or IsEmpty(Data(i, j))) Then
29                LastRow = i
30                LastCol = j
31                Exit For
32            End If
33        Next k

          Dim LastRow2 As Long, LastCol2 As Long
34        LastRow2 = LastRow: LastCol2 = LastCol
          
35        For j = NC To (LastCol + 1) Step -1
36            For i = 1 To LastRow
37                If Not (IsError(Data(i, j)) Or IsEmpty(Data(i, j))) Then
38                    LastCol2 = j
39                    GoTo EndLoop1
40                End If
41            Next i
42        Next j
EndLoop1:
43        For i = NR To LastRow - 1 Step -1
44            For j = 1 To LastCol
45                If Not (IsError(Data(i, j)) Or IsEmpty(Data(i, j))) Then
46                    LastRow2 = i
47                    GoTo EndLoop2
48                End If
49            Next j
50        Next i
EndLoop2:

51        If LastRow2 = NR And LastCol2 = NC Then
52            sTrimNAs = Data
53        Else
54            sTrimNAs = sSubArray(Data, 1, 1, LastRow2, LastCol2)
55        End If

56        Exit Function
ErrHandler:
57        sTrimNAs = "#sTrimNAs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
