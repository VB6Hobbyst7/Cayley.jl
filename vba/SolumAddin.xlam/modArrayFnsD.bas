Attribute VB_Name = "modArrayFnsD"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCompareTwoArrays
' Author    : Philip Swannell
' Date      : 27-May-2015
' Purpose   : Compares two arrays, and shows rows in common to both; rows in Array1 not in Array2; and
'             rows in Array2 not in Array1. Returns are sorted with duplicates removed and
'             there is a header row, unless the ControlString word "NoHeaders" is passed.
' Arguments
' Array1    : The first array of values of any type
' Array2    : The second array of values of any type. Must have the same number of columns as Array1.
' ControlString: Comma-separated with allowed words: "Common" or "C", "In1AndNotIn2" or "12",
'             "In2AndNotIn1" or "21". If omitted, all 3 are shown. Also: "CaseSens" for
'             case sensitive string comparison and "NoHeaders" to omit the header row.
' -----------------------------------------------------------------------------------------------------------------------
Function sCompareTwoArrays(ByVal Array1, ByVal Array2, Optional ByVal ControlString As String = "Common,In1AndNotIn2,In2AndNotIn1", Optional CompareRowsWithRows As Boolean = True)
Attribute sCompareTwoArrays.VB_Description = "Compares two arrays, and shows rows in common to both; rows in Array1 not in Array2; and rows in Array2 not in Array1. Returns are sorted with duplicates removed and there is a header row, unless the ControlString word ""NoHeaders"" is passed."
Attribute sCompareTwoArrays.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim CaseSensitive As Boolean
          Dim Choose1n2 As Variant
          Dim Choose2n1 As Variant
          Dim ChooseCommon1 As Variant
          Dim CleanControlString As String
          Dim ControlStringArray As Variant
          Dim Do1N2 As Boolean
          Dim Do2N1 As Boolean
          Dim DoCommon As Boolean
          Dim DoHeaders As Boolean
          Dim header As Variant
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim N1n2 As Long
          Dim N2n1 As Long
          Dim NC As Long
          Dim NCommon As Long
          Dim NR1 As Long
          Dim NR2 As Long
          Dim Result As Variant
          Dim Result1n2 As Variant
          Dim Result2n1 As Variant
          Dim ResultCommon As Variant

1         On Error GoTo ErrHandler

2         ControlString = Replace(ControlString, " ", vbNullString)        'Ignore spaces in ControlString words"
3         ControlStringArray = sTokeniseString(ControlString)

4         DoHeaders = True
5         For i = 1 To sNRows(ControlStringArray)
6             Select Case LCase$(ControlStringArray(i, 1))
                  Case "in1andnotin2", "12", "in1not2"
7                     ControlStringArray(i, 1) = "In1AndNotIn2"
8                     CleanControlString = CleanControlString & "In1AndNotIn2"
9                     Do1N2 = True
10                Case "in2andnotin1", "21", "in2not1"
11                    ControlStringArray(i, 1) = "In2AndNotIn1"
12                    Do2N1 = True
13                    CleanControlString = CleanControlString & "In2AndNotIn1"
14                Case "common", "c"
15                    ControlStringArray(i, 1) = "Common"
16                    DoCommon = True
17                    CleanControlString = CleanControlString & "Common"
18                Case "casesens", "casesensitive"
19                    CaseSensitive = True
20                Case "noheaders"
21                    DoHeaders = False
22                Case Else
23                    Throw "Unrecognised ControlString word. ControlString must be comma delimited with allowed words In1AndNotIn2, In2AndNotIn1, Common, CaseSens and NoHeaders"
24            End Select
25        Next i

26        If Not (DoCommon Or Do1N2 Or Do2N1) Then Throw "ControlString must contain at least one of Common, In1AndNotIn2 and In2AndNotIn1. Other allowed ControlString words are CaseSensitive, and NoHeaders"

27        Force2DArrayRMulti Array1, Array2
28        If Not CompareRowsWithRows Then
29            If sNCols(Array1) > 1 Then Array1 = sReshape(Array1, sNRows(Array1) * sNRows(Array1), 1)
30            If sNCols(Array2) > 1 Then Array2 = sReshape(Array2, sNRows(Array2) * sNRows(Array2), 1)
31        End If

32        NC = sNCols(Array1)
33        If sNCols(Array2) <> NC Then Throw "Array1 and Array2 must have the same number of columns, or try passing CompareRowsWithRows as FALSE"

34        Array1 = sRemoveDuplicateRows(Array1, True, CaseSensitive)
35        Array2 = sRemoveDuplicateRows(Array2, True, CaseSensitive)
36        NR1 = sNRows(Array1): NR2 = sNRows(Array2)

37        MatchIDs = sMultiMatchSortChop(Array1, Array2, NR1, NR2, NC, CaseSensitive, True)
38        Force2DArray MatchIDs
39        ChooseCommon1 = sArrayIsNumber(MatchIDs)
40        NCommon = sArrayCount(ChooseCommon1)
41        Choose1n2 = sArrayNot(ChooseCommon1)
42        N1n2 = NR1 - NCommon
43        Choose2n1 = sReshape(True, NR2, 1)
44        For i = 1 To NR1
45            If VarType(MatchIDs(i, 1)) <> vbString Then
46                Choose2n1(MatchIDs(i, 1), 1) = False
47            End If
48        Next i
49        N2n1 = NR2 - NCommon

          'Construct the chunks of the return array, have to handle 4 cases for each chunk
50        If DoCommon Then
51            If DoHeaders Then
52                header = sReshape(vbNullString, 1, NC)
53                header(1, 1) = "Common"
54                If NCommon > 0 Then
55                    ResultCommon = sArrayStack(header, sMChoose(Array1, ChooseCommon1))
56                Else
57                    ResultCommon = header
58                End If
59            Else
60                If NCommon > 0 Then
61                    ResultCommon = sMChoose(Array1, ChooseCommon1)
62                Else
63                    ResultCommon = sReshape(CVErr(xlErrNA), 1, NC)
64                End If
65            End If
66        End If
67        If Do1N2 Then
68            If DoHeaders Then
69                header = sReshape(vbNullString, 1, NC)
70                header(1, 1) = "In1AndNotIn2"
71                If N1n2 > 0 Then
72                    Result1n2 = sArrayStack(header, sMChoose(Array1, Choose1n2))
73                Else
74                    Result1n2 = header
75                End If
76            Else
77                If N1n2 > 0 Then
78                    Result1n2 = sMChoose(Array1, Choose1n2)
79                Else
80                    Result1n2 = sReshape(CVErr(xlErrNA), 1, NC)
81                End If
82            End If
83        End If
84        If Do2N1 Then
85            If DoHeaders Then
86                header = sReshape(vbNullString, 1, NC)
87                header(1, 1) = "In2AndNotIn1"
88                If N2n1 > 0 Then
89                    Result2n1 = sArrayStack(header, sMChoose(Array2, Choose2n1))
90                Else
91                    Result2n1 = header
92                End If
93            Else
94                If N2n1 > 0 Then
95                    Result2n1 = sMChoose(Array2, Choose2n1)
96                Else
97                    Result2n1 = sReshape(CVErr(xlErrNA), 1, NC)
98                End If
99            End If
100       End If

          'Put the chunks together in the final array. For the common case of the default control string _
           it's a bit faster to do a single call to sArrayRange
101       If CleanControlString = "CommonIn1AndNotIn2In2AndNotIn1" Then
102           sCompareTwoArrays = sArrayRange(ResultCommon, Result1n2, Result2n1)
103           Exit Function
104       Else
105           j = 0
106           For i = 1 To sNRows(ControlStringArray)
107               Select Case LCase$(ControlStringArray(i, 1))
                      Case "common"
108                       j = j + 1
109                       If j = 1 Then
110                           Result = ResultCommon
111                       Else
112                           Result = sArrayRange(Result, ResultCommon)
113                       End If
114                   Case "in1andnotin2"
115                       j = j + 1
116                       If j = 1 Then
117                           Result = Result1n2
118                       Else
119                           Result = sArrayRange(Result, Result1n2)
120                       End If
121                   Case "in2andnotin1"
122                       j = j + 1
123                       If j = 1 Then
124                           Result = Result2n1
125                       Else
126                           Result = sArrayRange(Result, Result2n1)
127                       End If
128               End Select
129           Next i
130           sCompareTwoArrays = Result
131       End If

132       Exit Function
ErrHandler:
133       sCompareTwoArrays = "#sCompareTwoArrays (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDifference
' Author    : Philip Swannell
' Date      : 29-Nov-2016
' Purpose   : Returns lagged and iterated differences. If order is 1 then Return(i,j) =
'             TheArray(i+Lag,j) - TheArray(i,j). If order > 1 then the algorithm is applied
'             recursively. The number of rows in the return is the number of rows in
'             TheArray minus Lag * Order.
' Arguments
' TheArray  : An arbitrary array of numbers.
' Lag       : The Lag. Must be at least one. If omitted defaults to one. Non-whole numbers are rounded
'             to the nearest whole number.
' Order     : The Order. Must be at least one. If omitted defaults to one. Non-whole numbers are rounded
'             to the nearest whole number.
'
' Notes     : See also sFirstDifference. Note that sFirstDifference(TheArray) =
'             sDifference(TheArray,1,1)
' -----------------------------------------------------------------------------------------------------------------------
Function sDifference(TheArray, Optional Lag As Long = 1, Optional order As Long = 1)
Attribute sDifference.VB_Description = "Returns lagged and iterated differences. If order is 1 then Return(i,j) = TheArray(i+Lag,j) - TheArray(i,j). If order > 1 then the algorithm is applied recursively. The number of rows in the return is the number of rows in TheArray minus Lag * Order."
Attribute sDifference.VB_ProcData.VB_Invoke_Func = " \n30"

1         On Error GoTo ErrHandler

          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result() As Variant
2         Force2DArrayR TheArray, N, M
3         If order < 0 Then Throw "Order must be greater than or equal to zero"
4         If Lag < 1 Then Throw "Lag must be one or more than one"

5         If N < ((Lag * order) + 1) Then Throw "Too few rows in TheArray"

6         If order = 0 Then
7             sDifference = TheArray
8             Exit Function
9         End If

10        ReDim Result(1 To N - Lag, 1 To M)
11        For i = 1 To N - Lag
12            For j = 1 To M
13                Result(i, j) = SafeSubtract(TheArray(i + Lag, j), TheArray(i, j))
14            Next j
15        Next i
16        sDifference = Result

17        If order = 1 Then
18            sDifference = Result
19        Else
20            sDifference = sDifference(Result, Lag, order - 1)
21        End If

22        Exit Function
ErrHandler:
23        Throw "#sDifference (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDiffTwoArrays
' Author    : Philip Swannell
' Date      : 04-Jan-2016
' Purpose   : See the difference between two arrays of arbitrary data. If elements of Array1 and Array2
'             are:
'             Both numbers, returns Array1 - Array2; else
'             Both the same, returns Array1; else
'             Returns a string listing both elements: [Array1Element,Array2Element]
' Arguments
' Array1    : An array of arbitrary data
' Array2    : An array of arbitrary data
' -----------------------------------------------------------------------------------------------------------------------
Function sDiffTwoArrays(ByVal Array1, ByVal Array2)
Attribute sDiffTwoArrays.VB_Description = "See the difference between two arrays of arbitrary data. If elements of Array1 and Array2 are:\nBoth numbers, returns Array1 - Array2; else\nBoth the same, returns Array1; else\nReturns a string listing both elements: [Array1Element,Array2Element]"
Attribute sDiffTwoArrays.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim C1 As Long
          Dim C2 As Long
          Dim CMax As Long
          Dim CMin As Long
          Dim i As Long
          Dim j As Long
          Dim r1 As Long
          Dim r2 As Long
          Dim Result() As Variant
          Dim RMax As Long
          Dim RMin As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti Array1, Array2

3         r1 = sNRows(Array1): C1 = sNCols(Array1)
4         r2 = sNRows(Array2): C2 = sNCols(Array2)
5         RMin = SafeMin(r1, r2): RMax = SafeMax(r1, r2)
6         CMin = SafeMin(C1, C2): CMax = SafeMax(C1, C2)

7         ReDim Result(1 To RMax, 1 To CMax)

8         For i = 1 To RMin
9             For j = 1 To CMin
10                If IsNumberOrDate(Array1(i, j)) And IsNumberOrDate(Array2(i, j)) Then
11                    Result(i, j) = SafeSubtract(Array1(i, j), Array2(i, j))
12                    If VarType(Result(i, j)) = vbDate Then
                          'Because the difference of dates is not a date!
13                        Result(i, j) = CDbl(Result(i, j))
14                    End If
15                ElseIf sEquals(Array1(i, j), Array2(i, j)) Then
16                    Result(i, j) = Array1(i, j)
17                    If IsEmpty(Array1(i, j)) Then Result(i, j) = vbNullString        ' to avoid Empties getting "cast" to zero when a UDF returns values to the sheet
18                Else
19                    Result(i, j) = "[" + NonStringToString(Array1(i, j), True) + "," + NonStringToString(Array2(i, j), True) + "]"
20                End If
21            Next j
22            For j = CMin + 1 To CMax
23                If C1 > C2 Then
24                    Result(i, j) = "[" + NonStringToString(Array1(i, j), True) + ",]"
25                Else
26                    Result(i, j) = "[," + NonStringToString(Array2(i, j), True) + "]"
27                End If
28            Next j
29        Next i
30        For i = RMin + 1 To RMax
31            For j = 1 To CMin
32                If r1 > r2 Then
33                    Result(i, j) = "[" + NonStringToString(Array1(i, j), True) + ",]"
34                Else
35                    Result(i, j) = "[," + NonStringToString(Array2(i, j), True) + "]"
36                End If
37            Next j
38            For j = CMin + 1 To CMax
                  'Result(i, j) = vbNullString
39            Next j
40        Next i

41        sDiffTwoArrays = Result
42        Exit Function
ErrHandler:
43        sDiffTwoArrays = "#sDiffTwoArrays (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sEveryNthElement
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : Returns every Nth row of the input array, starting at FirstRowNo.
' Arguments
' TheArray  : An array of arbitrary values.
' FirstRowNo: The index of the first row of the input array that is to be included in the output. Row
'             counting starts from 1.
' N         : The "step size", e.g. if N is 2 then every other row of the input array appears in the
'             output.
' Inverse   : Optional, defaulting to FALSE. If TRUE, then the function inverts its behaviour, inserting
'             rows of null-strings between the rows of the input TheArray. Thus the output
'             is an array whose "EveryNthElement" is the input TheArray.
' -----------------------------------------------------------------------------------------------------------------------
Function sEveryNthElement(ByVal TheArray As Variant, FirstRowNo As Long, N As Long, Optional Inverse As Boolean)
Attribute sEveryNthElement.VB_Description = "Returns every Nth row of the input array, starting at FirstRowNo."
Attribute sEveryNthElement.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NCols As Long
          Dim NRows As Long
          Dim NRowsOut As Long
          Dim Result() As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, NRows, NCols

3         If Not Inverse Then

4             If FirstRowNo < 1 Or FirstRowNo > NRows Then Throw "FirstRowNo must be in the range 1 to " + CStr(NRows)
5             If N < 1 Then Throw "N must be a positive number"

6             NRowsOut = (NRows - FirstRowNo) \ N + 1

7             ReDim Result(1 To NRowsOut, 1 To NCols)

8             For j = 1 To NCols
9                 k = FirstRowNo
10                For i = 1 To NRowsOut
11                    Result(i, j) = TheArray(k, j)
12                    k = k + N
13                Next i
14            Next j
15            sEveryNthElement = Result
16        Else
17            If FirstRowNo < 1 Then Throw "FirstRowNo must be greater than or equal to 1"
18            If N < 1 Then Throw "N must be a positive number"
19            NRowsOut = FirstRowNo + (NRows - 1) * N
20            ReDim Result(1 To NRowsOut, 1 To NCols)
21            For j = 1 To NCols
22                For i = 1 To NRowsOut
23                    Result(i, j) = vbNullString
24                Next
25            Next
26            For j = 1 To NCols
27                k = FirstRowNo
28                For i = 1 To NRows
29                    Result(k, j) = TheArray(i, j)
30                    k = k + N
31                Next i
32            Next j
33            sEveryNthElement = Result
34        End If

35        Exit Function
ErrHandler:
36        sEveryNthElement = "#sEveryNthElement (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileRegExReplace
' Author    : Philip Swannell
' Date      : 13-Jan-2016
' Purpose   : Uses regular expressions to make replacement within a text file. The function creates a
'             TargetFile whose contents is the contents of the SourceFile but with every
'             instance of the regular expression match replaced with the replacement.
' Arguments
' SourceFile: Full name (with path) of the source file.
' TargetFile: Full name (with path) of the target file. TargetFile can be the same as SourceFile.
' RegularExpression: A standard regular expression string.
' Replacement: A replacement template for each match of the regular expression in the input string.
' CaseSensitive: TRUE for case-sensitive matching, FALSE for case-insensitive matching. This argument is
'             optional, defaulting to FALSE for case-insensitive matching.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileRegExReplace(SourceFile As Variant, TargetFile As Variant, RegularExpression As String, Replacement As String, Optional CaseSensitive As Boolean)
Attribute sFileRegExReplace.VB_Description = "Uses regular expressions to make replacement within a text file. The function creates a TargetFile whose contents is the contents of the SourceFile but with every instance of the regular expression match replaced with the replacement."
Attribute sFileRegExReplace.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileRegExReplace = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileRegExReplace = Broadcast2Args(FuncIdFileRegExReplace, SourceFile, TargetFile, RegularExpression, Replacement, CaseSensitive)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileTranspose
' Author    : Philip Swannell
' Date      : 03-May-2021
' Purpose   : Writes a file containing the transpose of the input file.
' Arguments
' InputFile : The name (with path) of the input text file. May be an array in which case output file
'             must be an array of the same size.
' OutputFile: The name (with path) of the output text file. May be an array, in which case inputfile
'             must be an array of the same size.
' Delimiter : The delimiter character(s).
' -----------------------------------------------------------------------------------------------------------------------
Function sFileTranspose(InputFile As Variant, OutputFile As Variant, Delimiter As String)
Attribute sFileTranspose.VB_Description = "Writes a file containing the transpose of the input file."
Attribute sFileTranspose.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileTranspose = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileTranspose = Broadcast2Args(FuncIDFileTranspose, InputFile, OutputFile, Delimiter)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileSearchWithin
' Author    : Philip Swannell
' Date      : 13-Jan-2016
' Purpose   : Searches the contents of a text file for a match to a given regular expression. Returns
'             TRUE if a match is found, FALSE if a match is not found, or an error string.
' Arguments
' FileName  : The full name of the file, including the path.
' RegularExpression: A standard regular expression string. See sIsRegMatch.
' CaseSensitive: TRUE for case-sensitive matching, FALSE for case-insensitive matching. This argument is
'             optional, defaulting to FALSE for case-insensitive matching.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSearchWithin(FileName As String, RegularExpression As String, Optional CaseSensitive As Boolean)
Attribute sFileSearchWithin.VB_Description = "Searches the contents of a text file for a match to a given regular expression. Returns TRUE if a match is found, FALSE if a match is not found, or an error string."
Attribute sFileSearchWithin.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim FileContents As String
          Dim FSO As Scripting.FileSystemObject
          Dim t1 As TextStream

1         On Error GoTo ErrHandler

2         If Not RegExSyntaxValid(RegularExpression) Then
3             sFileSearchWithin = "#Invalid syntax for RegularExpression!"
4             Exit Function
5         End If

6         Set FSO = New FileSystemObject
7         Set t1 = FSO.OpenTextFile(FileName, ForReading)
8         FileContents = t1.ReadAll
9         sFileSearchWithin = sIsRegMatch(RegularExpression, FileContents, CaseSensitive)
10        t1.Close
11        Set t1 = Nothing: Set FSO = Nothing

12        Exit Function
ErrHandler:
13        sFileSearchWithin = "#sFileSearchWithin (line " & CStr(Erl) + "): " & Err.Description & "!"
14        If Not t1 Is Nothing Then
15            t1.Close
16            Set t1 = Nothing: Set FSO = Nothing
17        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFirstDifference
' Author    : Philip Swannell
' Date      : 22-Jun-2016
' Purpose   : Each column of the return is the "first differences" of the corresponding column of
'             TheArray i.e. Output(i,j) = TheArray(i+1,j)-TheArray(i,j)
' Arguments
' TheArray  : An array of arbitrary values, when non-numbers appear in TheArray elements of the return
'             from the function will be error strings
'
' Notes     : The function is equivalent to (though a bit faster than)
'             sArraySubtract(sDrop(TheArray,1),sDrop(TheArray,-1))
'
'             See also sDifference. Note that sFirstDifference(TheArray) =
'             sDifference(TheArray,1,1)
' -----------------------------------------------------------------------------------------------------------------------
Function sFirstDifference(TheArray)
Attribute sFirstDifference.VB_Description = "Each column of the return is the ""first differences"" of the corresponding column of TheArray i.e. Output(i,j) = TheArray(i+1,j)-TheArray(i,j)"
Attribute sFirstDifference.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result() As Variant
2         Force2DArrayR TheArray
3         M = sNCols(TheArray): N = sNRows(TheArray) - 1
4         If N < 2 Then Throw "TheArray must have at least two rows"
5         ReDim Result(1 To N, 1 To M)
6         For i = 1 To N
7             For j = 1 To M
8                 Result(i, j) = SafeSubtract(TheArray(i + 1, j), TheArray(i, j))
9             Next j
10        Next i
11        sFirstDifference = Result
12        Exit Function
ErrHandler:
13        sFirstDifference = "#sFirstDifference (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFirstRatio
' Author    : Philip Swannell
' Date      : 22-Jun-2016
' Purpose   : Each column of the return is the "first ratio" of the corresponding column of TheArray
'             i.e. Output(i,j) = TheArray(i+1,j)/TheArray(i,j)
' Arguments
' TheArray  : An array of arbitrary values, when non-numbers appear in TheArray elements of the return
'             from the function will be error strings
'
' Notes     : The function is equivalent to (though a bit faster than)
'             sArrayDivide(sDrop(TheArray,1),sDrop(TheArray,-1))
' -----------------------------------------------------------------------------------------------------------------------
Function sFirstRatio(TheArray)
Attribute sFirstRatio.VB_Description = "Each column of the return is the ""first ratio"" of the corresponding column of TheArray i.e. Output(i,j) = TheArray(i+1,j)/TheArray(i,j)"
Attribute sFirstRatio.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result() As Variant
2         Force2DArrayR TheArray
3         M = sNCols(TheArray): N = sNRows(TheArray) - 1
4         If N < 1 Then Throw "TheArray must have at least two rows"
5         ReDim Result(1 To N, 1 To M)
6         For i = 1 To N
7             For j = 1 To M
8                 Result(i, j) = SafeDivide(TheArray(i + 1, j), TheArray(i, j))
9             Next j
10        Next i
11        sFirstRatio = Result
12        Exit Function
ErrHandler:
13        sFirstRatio = "#sFirstRatio (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sGrid
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Returns a column of equally-spaced numbers with specified first, last and number of
'             elements.
' Arguments
' FirstElement: The top element in the column.
' LastElement: The bottom element in the column.
' NumElements: The height of the columm.
' -----------------------------------------------------------------------------------------------------------------------
Function sGrid(FirstElement As Double, LastElement As Double, NumElements As Long)
Attribute sGrid.VB_Description = "Returns a column of equally-spaced numbers with specified first, last and number of elements."
Attribute sGrid.VB_ProcData.VB_Invoke_Func = " \n30"

1         On Error GoTo ErrHandler
          Dim i As Long
          Dim ReturnArray() As Double

2         If NumElements < 0 Then
3             sGrid = "#NumElements must positive!"
4             Exit Function
5         ElseIf NumElements = 1 Then
6             If FirstElement <> LastElement Then
7                 sGrid = "When NumElements is 1, FirstElement and LastElement must be equal"
8             Else
9                 ReDim ReturnArray(1 To NumElements, 1 To 1)
10                ReturnArray(1, 1) = FirstElement
11                sGrid = ReturnArray
12                Exit Function
13            End If
14        End If

15        ReDim ReturnArray(1 To NumElements, 1 To 1)

16        For i = 1 To NumElements
17            ReturnArray(i, 1) = (FirstElement * (NumElements - i) + LastElement * (i - 1)) / (NumElements - 1)
18        Next i

19        sGrid = ReturnArray
20        Exit Function
ErrHandler:
21        sGrid = "#sGrid: line(" + CStr(Erl) + ") " + Err.Description + "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sGroupReshape
' Author    : Philip Swannell
' Date      : 02-May-2015
' Purpose   : The return consists of the first row of TheArray repeated NumCopies times, followed by the
'             second row of the input repeated NumCopies times etc until all rows of
'             TheArray have been represented NumCopies times.
' Arguments
' TheArray  : An array of arbitrary values.
' NumCopies : A positive integer giving how many repeats of the elements of TheArray are to be given.
' -----------------------------------------------------------------------------------------------------------------------
Function sGroupReshape(ByVal TheArray As Variant, NumCopies As Long)
Attribute sGroupReshape.VB_Description = "The return consists of the first row of TheArray repeated NumCopies times, followed by the second row of the input repeated NumCopies times etc until all rows of TheArray have been represented NumCopies times."
Attribute sGroupReshape.VB_ProcData.VB_Invoke_Func = " \n27"
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim L As Long
          Dim M As Long
          Dim N As Long
          Dim Result() As Variant
2         Force2DArrayR TheArray, N, M

3         If NumCopies < 1 Then Throw "NumCopies must be positive"
4         If NumCopies * N > 1048576 Then Throw "Maximum array size exceeded"
5         ReDim Result(1 To N * NumCopies, 1 To M)

6         For j = 1 To M
7             L = 0
8             For i = 1 To N
9                 For k = 1 To NumCopies
10                    L = L + 1
11                    Result(L, j) = TheArray(i, j)
12                Next k
13            Next i
14        Next j

15        sGroupReshape = Result

16        Exit Function
ErrHandler:
17        sGroupReshape = "#sGroupReshape (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sHistogram
' Author    : Philip Swannell
' Date      : 09-Jul-2017
' Purpose   : Generates data required for a histogram plot. Two columns are returned giving bucket
'             bounds and the count of elements within Data that lie in each bucket.
' Arguments
' Data      : An array of arbitrary values. Non numeric entries are ignored.
' Min       : The "bottom edge" of the bottom bucket. This argument is optional, if omitted it is chosen
'             so that the minimum of Data lies at the centre point of the bottom bucket.
' Max       : The "top edge" of the top bucket. This argument is optional, if omitted it is chosen so
'             that the maximum of Data lies at the centre point of the top bucket.
' NumBuckets: The number of buckets. This argument is optional and defaults to 10. It must be at least
'             2.
' CountOutliers: If TRUE, then the return also shows the counts of the number of elements in Data less than
'             or equal to Min, and the number of elements greater than Max. This argument
'             is optional and defaults to TRUE if either Max or Min are provided, and FALSE
'             otherwise.
' OneSeriesPerCol: Specifies whether multiple columns in Data represent separate data series (in which case
'             the function return has multiple columns) or are all part of the same data
'             series.
' SeriesNames: An optional array of the names of the data series. Used to construct the header row of the
'             return.
' NumberFormat: For construction of the labels in the left column of the return. Example: "0.00". When
'             NumberFormat is omitted and the function is being used on a worksheet, the
'             NumberFormat property of the left column of cells where the formula is
'             entered is used.
'
' Notes     : Example
'             If Data is a range with 2 columns and 1000 rows of standard normal deviates
'             then
'             {=sHistogram(Data,-3,3,6,TRUE,TRUE)} might return:
'
'             x            Series 1     Series 2
'             (-Inf, -3]        1            3
'             (-3, -2]         27           21
'             (-2, -1]        123          136
'             (-1, 0]         372          353
'             (0, 1]          331          324
'             (1, 2]          124          144
'             (2, 3]           22           17
'             (3, Inf)          0            2
' -----------------------------------------------------------------------------------------------------------------------
Function sHistogram(ByVal Data, Optional Min As Variant, Optional Max As Variant, Optional NumBuckets As Long = 10, Optional ByVal CountOutliers As Variant, _
        Optional OneSeriesPerCol As Boolean = True, Optional ByVal SeriesNames As Variant, Optional ByVal NumberFormat As String = vbNullString)
Attribute sHistogram.VB_Description = "Generates data required for a histogram or density plot. Two columns are returned giving bucket bounds and the count of elements within Data that lie in each bucket."
Attribute sHistogram.VB_ProcData.VB_Invoke_Func = " \n27"
        
          Dim Boundaries
          Dim colFrequency As Variant
          Dim dMax As Double
          Dim dMin As Double
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim BucketWidth As Double
          Dim LeftCol As Variant
          Dim RequiredNumElements As Long
          Dim tn As String

1         On Error GoTo ErrHandler
          'If NumberFormat not supplied then if possible grab it from the range into which the formula is entered. Is this too clever by half?
2         If NumberFormat = vbNullString Then
3             If TypeName(Application.Caller) = "Range" Then
4                 NumberFormat = "General"
5                 On Error Resume Next
6                 NumberFormat = Application.Caller.Cells(2, 1).NumberFormat
7                 On Error GoTo ErrHandler
8             Else
9                 NumberFormat = "General"
10            End If
11        End If

12        If IsMissing(CountOutliers) Or IsEmpty(CountOutliers) Then
13            CountOutliers = Not (IsMissing(Max) And IsMissing(Min))
14        ElseIf VarType(CountOutliers) <> vbBoolean Then
15            Throw "CountOutliers must be TRUE, FALSE or omitted"
16        End If
17        If Not IsMissing(Min) Then
18            If Not IsEmpty(Min) Then
19                If Not IsNumberOrDate(Min) Then
20                    If TypeName(Min) = "Range" Then
21                        tn = TypeName(Min.Value)
22                    Else
23                        tn = TypeName(Min)
24                    End If
25                    Throw "Min must be a number or missing or empty, but it is a " + LCase$(tn)
26                End If
27            End If
28        End If
29        If Not IsMissing(Max) Then
30            If Not IsEmpty(Max) Then
31                If Not IsNumberOrDate(Max) Then
32                    If TypeName(Max) = "Range" Then
33                        tn = TypeName(Max.Value)
34                    Else
35                        tn = TypeName(Max)
36                    End If
37                    Throw "Max must be a number or missing or empty, but it is a " + LCase$(tn)
38                End If
39            End If
40        End If

41        Force2DArrayR Data, NR, NC
          'Needed since Application.Worksheet.Frequency is happy to ignore strings, but not to ignore error values
42        For j = 1 To NC
43            For i = 1 To NR
44                If IsError(Data(i, j)) Then
45                    Data(i, j) = CStr(Data(i, j))
46                End If
47            Next i
48        Next j

49        If NumBuckets < 2 Then Throw "NumBuckets must be at least 2"
50        If (Not IsNumberOrDate(Max)) And (Not IsNumberOrDate(Min)) Then
51            dMax = ThrowIfError(sMaxOfNums(Data))
52            dMin = ThrowIfError(sMinOfNums(Data))
53            If dMax = dMin Then
54                If dMax = 0 Then
55                    dMax = 1
56                    dMin = -1
57                Else
58                    dMax = 1.1 * dMax
59                    dMin = 0.9 * dMin
60                End If
61            Else
                  'Make the top bucket contain the maximum value at its centre point _
                   and the bottom bucket contain the minimum value at its centre point
62                BucketWidth = (dMax - dMin) / (NumBuckets - 1)
63                dMax = dMax + BucketWidth / 2
64                dMin = dMin - BucketWidth / 2
65            End If
66        ElseIf Not IsNumberOrDate(Max) Then
              'Make the top bucket contain the maximum value at its centre point
67            dMax = ThrowIfError(sMaxOfNums(Data))
68            dMin = Min
69            BucketWidth = (dMax - dMin) / (NumBuckets - 0.5)
70            dMax = dMax + BucketWidth / 2
71        ElseIf Not IsNumberOrDate(Min) Then
              'Make the bottom bucket contain the minimum value at its centre point
72            dMax = Max
73            dMin = ThrowIfError(sMinOfNums(Data))
74            BucketWidth = (dMax - dMin) / (NumBuckets - 0.5)
75            dMin = dMin - BucketWidth / 2
76        Else
77            dMax = Max
78            dMin = Min
79        End If

80        If dMin >= dMax Then Throw "Max must be greater than Min"
81        Boundaries = ThrowIfError(sGrid(dMin, dMax, NumBuckets + 1))

82        NC = sNCols(Data): NR = sNRows(Data)
83        If OneSeriesPerCol Then
84            colFrequency = sReshape(0, NumBuckets + 2, NC)
85        Else
86            colFrequency = sReshape(0, NumBuckets + 2, 1)
87        End If

88        If IsMissing(SeriesNames) Then
89            If OneSeriesPerCol Then
                  Dim tmp() As String
90                ReDim tmp(1 To 1, 1 To NC)
91                For i = 1 To NC
92                    tmp(1, i) = "Series " & CStr(i)
93                Next
94                SeriesNames = tmp
95            Else
96                SeriesNames = "Frequency"
97            End If
98        Else
99            Force2DArrayR SeriesNames
100           RequiredNumElements = IIf(OneSeriesPerCol, NC, 1)
101           If (sNRows(SeriesNames) * sNCols(SeriesNames)) <> RequiredNumElements Then
102               Throw "SeriesNames must be omitted or provided with " & CStr(RequiredNumElements) & " element(s)"
103           End If
104           If sNRows(SeriesNames) <> 1 Then
105               SeriesNames = sReshape(SeriesNames, 1, RequiredNumElements)
106           End If
107       End If

108       If OneSeriesPerCol Then
109           colFrequency = sReshape(0, NumBuckets + 2, NC)
              Dim FRRes
110           For j = 1 To NC
111               FRRes = Application.WorksheetFunction.Frequency(sSubArray(Data, 1, j, , 1), Boundaries)
112               For i = 1 To sNRows(FRRes)
113                   colFrequency(i, j) = FRRes(i, 1)
114               Next i
115           Next j
116       Else
117           colFrequency = Application.WorksheetFunction.Frequency(Data, Boundaries)
118       End If
          
119       LeftCol = sHistogramLabels(dMin, dMax, NumBuckets, CBool(CountOutliers), CStr(NumberFormat))
120       If CountOutliers Then
121           sHistogram = sArraySquare("x", SeriesNames, LeftCol, colFrequency)
122       Else
123           sHistogram = sArraySquare("x", SeriesNames, LeftCol, sSubArray(colFrequency, 2, 1, NumBuckets))
124       End If

125       Exit Function
ErrHandler:
126       sHistogram = "#sHistogram (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sHistogramLabels
' Author     : Philip Swannell
' Date       : 22-Nov-2018
' Purpose    : Sub-routine of sHistogram, constructs "friendly" labels for the left column
' -----------------------------------------------------------------------------------------------------------------------
Private Function sHistogramLabels(Min As Double, Max As Double, NumBuckets As Long, CountOutliers As Boolean, NumberFormat As String)
          
          Dim c As Collection
          Dim dBoundaries
          Dim EN As Long
          Dim i As Long
          Dim LeftCol
          Dim Res As String
          Dim sBoundaries
          Dim tmp As String
1         On Error GoTo ErrHandler

2         If LCase$(NumberFormat) <> "density" Then
3             Res = "Foo"
4             On Error Resume Next
5             Res = Application.WorksheetFunction.text(0, NumberFormat)
6             On Error GoTo ErrHandler
7             If Res = "Foo" Then Throw "NumberFormat not recognised"
8         Else
              Dim BucketWidth As Double
              Dim FirstElement As Double
              Dim LastElement As Double
              Dim NumElements As Long
9             BucketWidth = (Max - Min) / NumBuckets
10            If CountOutliers Then
11                FirstElement = Min - BucketWidth / 2
12                LastElement = Max + BucketWidth / 2
13                NumElements = NumBuckets + 2
14            Else
15                FirstElement = Min + BucketWidth / 2
16                LastElement = Max - BucketWidth / 2
17                NumElements = NumBuckets
18            End If
19            sHistogramLabels = sGrid(FirstElement, LastElement, NumElements)
20            Exit Function
21        End If

22        dBoundaries = sGrid(Min, Max, NumBuckets + 1)
23        sBoundaries = sReshape(vbNullString, NumBuckets + 1, 1)
24        Set c = New Collection
25        For i = 1 To NumBuckets + 1
26            tmp = Application.WorksheetFunction.text(dBoundaries(i, 1), NumberFormat)
27            On Error Resume Next
28            c.Add 1, tmp
29            EN = Err.Number
30            On Error GoTo ErrHandler
31            If EN <> 0 Then Throw "NumberFormat '" + NumberFormat + "' does not make all " + CStr(NumBuckets + 1) + " bucket boundaries distinguishable. Try a NumberFormat with more significant figures."
32            sBoundaries(i, 1) = tmp
33        Next i

34        If CountOutliers Then
35            LeftCol = sReshape(vbNullString, NumBuckets + 2, 1)
36            LeftCol(1, 1) = "(-Inf, " & sBoundaries(1, 1) + "]"
37            For i = 1 To NumBuckets
38                LeftCol(1 + i, 1) = "(" & sBoundaries(i, 1) + ", " & sBoundaries(i + 1, 1) + "]"
39            Next
40            LeftCol(NumBuckets + 2, 1) = "(" & sBoundaries(NumBuckets + 1, 1) + ", Inf)"
41        Else
42            LeftCol = sReshape(vbNullString, NumBuckets, 1)
43            For i = 1 To NumBuckets
44                LeftCol(i, 1) = "(" & sBoundaries(i, 1) + ", " & sBoundaries(i + 1, 1) + "]"
45            Next
46        End If
47        sHistogramLabels = LeftCol
48        Exit Function
ErrHandler:
49        Throw "#sHistogramLabels (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
