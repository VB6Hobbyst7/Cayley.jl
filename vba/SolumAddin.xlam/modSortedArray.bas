Attribute VB_Name = "modSortedArray"
Option Explicit
Private m_ArrayToSort As Variant        'Used by sSortedArray, sSortedArray2 and sSortedArrayByAllCols
Private m_IgnoreFirstColumn As Variant  'Used by sSortedArrayByAllCols
Private m_NumCols As Long               'Used by sSortedArrayByAllCols
Private mColNo1 As Long                 'Used by sSortedArray
Private mColNo2 As Long                 'Used by sSortedArray
Private mColNo3 As Long                 'Used by sSortedArray
Private mAscending1 As Boolean          'Used by sSortedArray
Private mAscending2 As Boolean          'Used by sSortedArray
Private mAscending3 As Boolean          'Used by sSortedArray
Private mKeyCols                        'Used by sSortedArray2
Private mAscendings                     'Used by sSortedArray2
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSortedArray
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Returns the rows of an arbitrary array of values keyed on up to three columns. If two rows
'             match on all sort keys provided then their order is unchanged.
' Arguments
' ArrayToSort: An array of arbitrary values.
' ColNo1    : The column number for the first sort key, counting from left to right starting at 1. When
'             values in the first sort key column match then the method looks at the
'             second. If omitted defaults to 1 for the leftmost column.
' ColNo2    : The column number for the second sort key. When values in the first two sort key columns
'             match then the method looks at the third. If omitted no second sort key is
'             used.
' ColNo3    : The column number for the third sort key. When two rows match at all three sort key
'             columns then their relative order in the output will be unchanged. If omitted
'             no third sort key is used.
' Ascending1: TRUE if the output is to be in ascending order on the first sort key. If omitted defaults
'             to TRUE.
' Ascending2: TRUE if the output is to be in ascending order on the second sort key, FALSE for
'             descending. If omitted and ColNo2 is provided defaults to TRUE.
' Ascending3: TRUE if the output is to be in ascending order on the third sort key, FALSE for
'             descending. If omitted and ColNo3 is provided defaults to TRUE.
' CaseSensitive: Specifies if sorting of strings is case-sensitive.
'
'Notes:       Known differences with Excel's built-in sort (case-sensitive)
'          1) Empty values are sorted as being less than numbers, i.e. appear at the top of the return
'             in Ascending mode, and at the bottom in descending mode. In Excel's sort method
'             Emptys appear at the bottom of the sorted list both in ascending and descending modes.
'          2) Error values are sorted as described above, the sorting behaviour for error values
'             in Excel's sort method are difficult to understand.
'          3) Handling of high ascii characters is strange in Excel's built in sorting - characters
'             seem to be treated differently when they are the first character of a string from when
'             they are second or later character, so that sorting is not lexicographic.
'             Example Œ is chr(140) and we have o < Œ but oo > ŒŒ
'             No attempt to emulate this behaviour...
'TODO make handle 0-based arrays!
' -----------------------------------------------------------------------------------------------------------------------
Public Function sSortedArray(ArrayToSort As Variant, _
        Optional ColNo1 As Long = 1, _
        Optional ByVal ColNo2 As Long, _
        Optional ByVal ColNo3 As Long, _
        Optional Ascending1 As Boolean = True, _
        Optional Ascending2 As Boolean = True, _
        Optional Ascending3 As Boolean = True, _
        Optional CaseSensitive As Boolean = False, _
        Optional UseExcelSortMethod As Boolean = False, _
        Optional NumHeaderRows As Long)
Attribute sSortedArray.VB_Description = "Returns the rows of an arbitrary array of values keyed on up to three columns. If two rows match on all sort keys provided then their order is unchanged."
Attribute sSortedArray.VB_ProcData.VB_Invoke_Func = " \n27"

          'Check the inputs
1         On Error GoTo ErrHandler
          
2         If FunctionWizardActive() Then
3             sSortedArray = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         If TypeName(ArrayToSort) = "Range" Then
7             m_ArrayToSort = ArrayToSort.Value2
8         Else
9             m_ArrayToSort = ArrayToSort
10        End If
11        If NumHeaderRows < 0 Then
12            Throw "NumHeaderRows must be greater than or equal to zero"
13        ElseIf NumHeaderRows = 0 Then
14        ElseIf NumHeaderRows >= sNRows(ArrayToSort) Then
15            Throw "NumHeaderRows must be less than the number of rows in ArrayToSort"
16        End If

17        Select Case NumDimensions(m_ArrayToSort)
              Case 0
18                sSortedArray = m_ArrayToSort
19                Exit Function
20            Case 1, Is > 2
21                sSortedArray = "#ArrayToSort must be a two-dimensional array!"
22                Exit Function
23        End Select
24        If ColNo1 > UBound(m_ArrayToSort, 2) Or ColNo1 < 1 Then
25            sSortedArray = "#ColNo1 must be a valid column number!"
26            Exit Function
27        ElseIf ColNo2 > UBound(m_ArrayToSort, 2) Or ColNo2 < 0 Then
28            sSortedArray = "#ColNo2 must be a valid column number!"
29            Exit Function
30        ElseIf ColNo3 > UBound(m_ArrayToSort, 2) Or ColNo3 < 0 Then
31            sSortedArray = "#ColNo3 must be a valid column number!"
32            Exit Function
33        ElseIf ColNo2 = 0 And ColNo3 <> 0 Then
34            sSortedArray = "#ColNo2 must be provided when ColNo3 is provided!"
35            Exit Function
36        ElseIf (ColNo1 = ColNo2) Or (ColNo1 = ColNo3) Or (ColNo2 = ColNo3 And ColNo3 <> 0) Then
37            sSortedArray = "#Arguments ColNo1, ColNo2 and ColNo3 must be distict!"
38        End If

39        If UseExcelSortMethod Then
40            If ExcelSupportsSpill() Then
41                If CaseSensitive Then m_ArrayToSort = EncodeArray(m_ArrayToSort)
42                m_ArrayToSort = WrapSORT(m_ArrayToSort, ColNo1, ColNo2, ColNo3, Ascending1, Ascending2, Ascending3, NumHeaderRows)
43                If CaseSensitive Then m_ArrayToSort = DecodeArray(m_ArrayToSort)
44                sSortedArray = m_ArrayToSort
45                Exit Function
46            ElseIf TypeName(Application.Caller) <> "Range" Then
47                sSortedArray = SortWrap(m_ArrayToSort, ColNo1, ColNo2, ColNo3, Ascending1, Ascending2, Ascending3, CaseSensitive, NumHeaderRows)
48                Exit Function
49            Else
50                If ExcelSupportsSpill() Then
51                    Throw "Cannot do case-sensitive sort using Excel sort method when function is called from spreadsheet"
52                Else
53                    Throw "Cannot sort using Excel sort method when function is called from spreadsheet"
54                End If
55            End If
56        End If

          'Copy to module-level variables, so that we can access from CompareRows2
57        mColNo1 = ColNo1
58        mColNo2 = ColNo2
59        mColNo3 = ColNo3
60        mAscending1 = Ascending1
61        mAscending2 = Ascending2
62        mAscending3 = Ascending3

          Dim i As Long
          Dim NumCols As Long
          Dim NumRows As Long
          Dim vArray() As Long
63        NumRows = UBound(m_ArrayToSort, 1)
64        NumCols = UBound(m_ArrayToSort, 2)
65        ReDim vArray(1 To NumRows)
66        For i = 1 To NumRows
67            vArray(i) = i
68        Next

69        QuickSort2 vArray, NumHeaderRows + 1, UBound(vArray), CaseSensitive
          'vArray now acts as an idex for constructing ReturnArray
          Dim j As Long
          Dim ReturnArray() As Variant
70        ReDim ReturnArray(1 To UBound(m_ArrayToSort, 1), 1 To UBound(m_ArrayToSort, 2))
71        For i = 1 To NumRows
72            For j = 1 To NumCols
73                ReturnArray(i, j) = m_ArrayToSort(vArray(i), j)
74            Next j
75        Next i

76        sSortedArray = ReturnArray

77        Exit Function
ErrHandler:
78        sSortedArray = "#sSortedArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompareRows2
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Compares two rows of the module-level variable m_ArrayToSort, according to
'             module-level variables mColNo1, mColNo2, mColNo3, mAscending1, mAscending2, mAscending3
' -----------------------------------------------------------------------------------------------------------------------
Private Function CompareRows2(RowNo1 As Long, RowNo2 As Long, CaseSensitive As Boolean) As Boolean
          Dim EqualOnCol1 As Boolean
          Dim EqualOnCol2 As Boolean
          Dim EqualOnCol3 As Boolean
          Dim LessThanOnCol1 As Boolean
          Dim LessThanOnCol2 As Boolean
          Dim LessThanOnCol3 As Boolean

1         On Error GoTo ErrHandler

2         LessThanOnCol1 = VariantLessThan(m_ArrayToSort(RowNo1, mColNo1), m_ArrayToSort(RowNo2, mColNo1), CaseSensitive)
3         If LessThanOnCol1 Then
4             EqualOnCol1 = False
5         Else
6             EqualOnCol1 = sEquals(m_ArrayToSort(RowNo1, mColNo1), m_ArrayToSort(RowNo2, mColNo1), CaseSensitive)
7         End If

8         If LessThanOnCol1 Then
9             CompareRows2 = (mAscending1 = True)
10            Exit Function
11        ElseIf Not EqualOnCol1 Then
12            CompareRows2 = (mAscending1 = False)
13            Exit Function
14        Else
              'they match on col1, go on to look at the next column
15            If mColNo2 = 0 Then
16                CompareRows2 = RowNo1 < RowNo2        'Get a static sort this way
17                Exit Function
18            Else
19                LessThanOnCol2 = VariantLessThan(m_ArrayToSort(RowNo1, mColNo2), m_ArrayToSort(RowNo2, mColNo2), CaseSensitive)
20                If LessThanOnCol2 Then
21                    EqualOnCol2 = False
22                Else
23                    EqualOnCol2 = sEquals(m_ArrayToSort(RowNo1, mColNo2), m_ArrayToSort(RowNo2, mColNo2), CaseSensitive)
24                End If
25                If LessThanOnCol2 Then
26                    CompareRows2 = (mAscending2 = True)
27                    Exit Function
28                ElseIf Not EqualOnCol2 Then
29                    CompareRows2 = (mAscending2 = False)
30                    Exit Function
31                Else
                      'they match on col1 and col2, go on to look at the next col3
32                    If mColNo3 = 0 Then
33                        CompareRows2 = RowNo1 < RowNo2        'Get a static sort this way
34                        Exit Function
35                    Else
36                        LessThanOnCol3 = VariantLessThan(m_ArrayToSort(RowNo1, mColNo3), m_ArrayToSort(RowNo2, mColNo3), CaseSensitive)
37                        If LessThanOnCol3 Then
38                            EqualOnCol3 = False
39                        Else
40                            EqualOnCol3 = sEquals(m_ArrayToSort(RowNo1, mColNo3), m_ArrayToSort(RowNo2, mColNo3), CaseSensitive)
41                        End If
42                        If LessThanOnCol3 Then
43                            CompareRows2 = (mAscending3 = True)
44                            Exit Function
45                        ElseIf Not EqualOnCol3 Then
46                            CompareRows2 = (mAscending3 = False)
47                            Exit Function
48                        Else
49                            CompareRows2 = RowNo1 < RowNo2        'Get a static sort this way
50                        End If
51                    End If
52                End If
53            End If
54        End If
55        Exit Function
ErrHandler:
56        Throw "#CompareRows2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSort2
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Quick sort algorith taken from
'             http://stackoverflow.com/questions/152319/vba-array-sort-function
'             but changed so that at lines 6 and 9, where comparisons are done, rather than
'             comparing the contents of vArray we compare the (as yet unsorted) rows of m_ArrayToSort.
'             Hence ordering is inherited from method VariantLessThan, and in the case of strings
'             from function StringLessThan which is designed to mimic Excel's ordering of strings
'             implicit in DATA > Sort.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub QuickSort2(vArray() As Long, inLow As Long, inHi As Long, CaseSensitive As Boolean)

          Dim pivot As Long
          Dim tmpHi As Long
          Dim tmpLow As Long
          Dim tmpSwap As Variant

1         On Error GoTo ErrHandler

2         tmpLow = inLow
3         tmpHi = inHi

4         pivot = vArray((inLow + inHi) \ 2)

5         Do While (tmpLow <= tmpHi)

6             Do While (CompareRows2(vArray(tmpLow), pivot, CaseSensitive) And tmpLow < inHi)
7                 tmpLow = tmpLow + 1
8             Loop

9             Do While (CompareRows2(pivot, vArray(tmpHi), CaseSensitive) And tmpHi > inLow)
10                tmpHi = tmpHi - 1
11            Loop

12            If (tmpLow <= tmpHi) Then
13                tmpSwap = vArray(tmpLow)
14                vArray(tmpLow) = vArray(tmpHi)
15                vArray(tmpHi) = tmpSwap
16                tmpLow = tmpLow + 1
17                tmpHi = tmpHi - 1
18            End If

19        Loop

20        If (inLow < tmpHi) Then QuickSort2 vArray, inLow, tmpHi, CaseSensitive
21        If (tmpLow < inHi) Then QuickSort2 vArray, tmpLow, inHi, CaseSensitive

22        Exit Sub
ErrHandler:
23        Throw "#QuickSort2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestsSortedArray
' Author    : Philip Swannell
' Date      : 24-Nov-2016
' Purpose   : Test harness
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestsSortedArray()
          Dim k As Long
          Dim Randoms As Variant
          Dim Res1
          Dim Res2
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim TheArray

          Dim i As Long
          Dim j As Long
          Dim NA As Variant
          Dim NC As Long
          Dim NR As Long
1         NR = 10000
2         NC = 3

3         NA = CVErr(xlErrNA)

4         Randoms = (sRandomVariable(NR, NC, "Integer", , 26))
5         TheArray = sReshape(vbNullString, NR, NC)

6         For i = 1 To NR
7             For j = 1 To NC
8                 k = k + 1
9                 Select Case k Mod 4
                      Case 0, 1, 2, 3
10                        TheArray(i, j) = String(1, Chr$(96 + Randoms(i, j)))
11                    Case 2
12                        TheArray(i, j) = Randoms(i, j)
13                    Case 3
14                        TheArray(i, j) = Empty
15                End Select
16            Next
17        Next

18        t1 = sElapsedTime
19        Res1 = sSortedArray(TheArray, 1, 2, 3, False, True, False)
20        t2 = sElapsedTime
21        Res2 = WrapSORT(TheArray, 1, 2, 3, False, True, False)
22        t3 = sElapsedTime

23        Debug.Print sArraysIdentical(Res1, Res2), t2 - t1, t3 - t2, (t2 - t1) / (t3 - t2)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WrapSORT
' Author     : Philip Swannell
' Date       : 10-Dec-2019
' Purpose    : Wraps Excel's (new) SORT function, with signature as per sSortedArray
' -----------------------------------------------------------------------------------------------------------------------
Private Function WrapSORT(ByVal ArrayToSort As Variant, _
        Optional ColNo1 As Long = 1, _
        Optional ByVal ColNo2 As Long, _
        Optional ByVal ColNo3 As Long, _
        Optional Ascending1 As Boolean = True, _
        Optional Ascending2 As Boolean = True, _
        Optional Ascending3 As Boolean = True, _
        Optional NumHeaderRows As Long = 0)
          
          Dim Headers, NC As Long
          
1         On Error GoTo ErrHandler
2         If TypeName(ArrayToSort) = "Range" Then ArrayToSort = ArrayToSort.Value2

3         If NumHeaderRows > 0 Then
4             Headers = sSubArray(ArrayToSort, 1, 1, NumHeaderRows)
5             ArrayToSort = sSubArray(ArrayToSort, NumHeaderRows + 1)
6         End If
7         NC = sNCols(ArrayToSort)

          'either use SORT...
8         If ColNo2 = 0 And ColNo3 = 0 Then
9             ArrayToSort = Application.WorksheetFunction.Sort(ArrayToSort, ColNo1, IIf(Ascending1, 1, -1))
10            If NumHeaderRows > 0 Then
11                ArrayToSort = sArrayStack(Headers, ArrayToSort)
12            End If
13        Else
              'or use SORTBY...
              Dim Ascendings, KeyCols
14            Ascendings = CreateMissing()
15            KeyCols = CreateMissing()
16            If ColNo1 > 0 And ColNo1 <= NC Then
17                KeyCols = sArrayStack(KeyCols, ColNo1)
18                Ascendings = sArrayStack(Ascendings, Ascending1)
19            End If
20            If ColNo2 > 0 And ColNo2 <= NC Then
21                KeyCols = sArrayStack(KeyCols, ColNo2)
22                Ascendings = sArrayStack(Ascendings, Ascending2)
23            End If
24            If ColNo3 > 0 And ColNo3 <= NC Then
25                KeyCols = sArrayStack(KeyCols, ColNo3)
26                Ascendings = sArrayStack(Ascendings, Ascending3)
27            End If
28            ArrayToSort = WrapSORTBY(ArrayToSort, KeyCols, Ascendings, NumHeaderRows)
29        End If

30        Force2DArray ArrayToSort

31        WrapSORT = ArrayToSort

32        Exit Function
ErrHandler:
33        Throw "#WrapSORT (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WrapSORT
' Author     : Philip Swannell
' Date       : 10-Dec-2019
' Purpose    : Wraps Excel's (new) SORTBY function, with signature as per sSortedArray2
' -----------------------------------------------------------------------------------------------------------------------
Private Function WrapSORTBY(ByVal ArrayToSort, KeyCols, Ascendings, NumHeaderRows As Long)
          Dim Headers
1         On Error GoTo ErrHandler
2         If NumHeaderRows > 0 Then
3             Headers = sSubArray(ArrayToSort, 1, 1, NumHeaderRows)
4             ArrayToSort = sSubArray(ArrayToSort, NumHeaderRows + 1)
5         End If
6         Select Case sNRows(KeyCols)
              Case 1
7                 ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1))
8             Case 2
9                 ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1))
10            Case 3
11                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1))
12            Case 4
13                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1))
14            Case 5
15                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1))
16            Case 6
17                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1))
18            Case 7
19                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(7, 1), , 1), IIf(Ascendings(7, 1), 1, -1))
20            Case 8
21                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(7, 1), , 1), IIf(Ascendings(7, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(8, 1), , 1), IIf(Ascendings(8, 1), 1, -1))
22            Case 8
23                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(7, 1), , 1), IIf(Ascendings(7, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(8, 1), , 1), IIf(Ascendings(8, 1), 1, -1))
24            Case 9
25                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(7, 1), , 1), IIf(Ascendings(7, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(8, 1), , 1), IIf(Ascendings(8, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(9, 1), , 1), IIf(Ascendings(9, 1), 1, -1))
26            Case 10
27                ArrayToSort = Application.WorksheetFunction.SortBy(ArrayToSort, _
                      sSubArray(ArrayToSort, 1, KeyCols(1, 1), , 1), IIf(Ascendings(1, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(2, 1), , 1), IIf(Ascendings(2, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(3, 1), , 1), IIf(Ascendings(3, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(4, 1), , 1), IIf(Ascendings(4, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(5, 1), , 1), IIf(Ascendings(5, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(6, 1), , 1), IIf(Ascendings(6, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(7, 1), , 1), IIf(Ascendings(7, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(8, 1), , 1), IIf(Ascendings(8, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(9, 1), , 1), IIf(Ascendings(9, 1), 1, -1), _
                      sSubArray(ArrayToSort, 1, KeyCols(10, 1), , 1), IIf(Ascendings(10, 1), 1, -1))
28            Case Else
29                Throw "More than 10 elements in KeyCols is not supported"
30        End Select
31        If NumHeaderRows > 0 Then
32            ArrayToSort = sArrayStack(Headers, ArrayToSort)
33        End If
34        WrapSORTBY = ArrayToSort
35        Exit Function
ErrHandler:
36        Throw "#WrapSORTBY (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortWrap
' Author    : Philip Swannell
' Date      : 24-Nov-2016
' Purpose   : Version of sSortedArray that wraps Excel's .Sort method - cannot be called
'             from worksheet functions.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SortWrap(ByVal ArrayToSort As Variant, _
        Optional ColNo1 As Long = 1, _
        Optional ByVal ColNo2 As Long, _
        Optional ByVal ColNo3 As Long, _
        Optional Ascending1 As Boolean = True, _
        Optional Ascending2 As Boolean = True, _
        Optional Ascending3 As Boolean = True, _
        Optional CaseSensitive As Boolean = False, _
        Optional NumHeaderRows As Long = 0)

          Dim R As Range
          Dim Res
          Dim SUH As clsScreenUpdateHandler

1         Set SUH = CreateScreenUpdateHandler()

2         On Error GoTo ErrHandler
3         If TypeName(ArrayToSort) = "Range" Then ArrayToSort = ArrayToSort.Value2

4         With shEmptySheet
5             .Unprotect
6             .UsedRange.Clear
7             .UsedRange.EntireRow.Delete
8             Set R = .Cells(1, 1).Resize(sNRows(ArrayToSort), sNCols(ArrayToSort))
9             R.Value = sArrayExcelString(ArrayToSort)
10        End With

11        shEmptySheet.Sort.SortFields.Clear

12        If ColNo1 > 0 And ColNo1 <= R.Columns.Count Then
13            shEmptySheet.Sort.SortFields.Add Key:=R.Columns(ColNo1) _
                  , SortOn:=xlSortOnValues, order:=IIf(Ascending1, xlAscending, xlDescending), DataOption:=xlSortNormal
14        End If
15        If ColNo2 > 0 And ColNo2 <= R.Columns.Count Then
16            shEmptySheet.Sort.SortFields.Add Key:=R.Columns(ColNo2) _
                  , SortOn:=xlSortOnValues, order:=IIf(Ascending2, xlAscending, xlDescending), DataOption:=xlSortNormal
17        End If
18        If ColNo3 > 0 And ColNo3 <= R.Columns.Count Then
19            shEmptySheet.Sort.SortFields.Add Key:=R.Columns(ColNo3) _
                  , SortOn:=xlSortOnValues, order:=IIf(Ascending3, xlAscending, xlDescending), DataOption:=xlSortNormal
20        End If

21        With shEmptySheet.Sort
22            .SetRange R.Offset(NumHeaderRows).Resize(R.Rows.Count - NumHeaderRows)
23            .header = xlNo
24            .MatchCase = CaseSensitive
25            .Orientation = xlTopToBottom
26            .SortMethod = xlPinYin
27            .Apply
28        End With
29        SortWrap = R.Value
30        shEmptySheet.UsedRange.EntireRow.Delete
31        Res = shEmptySheet.UsedRange.Rows.Count

32        Exit Function
ErrHandler:
33        Throw "#SortWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortWrap2
' Author    : Philip Swannell
' Date      : 24-Nov-2016
' Purpose   : Version of sSortedArray that wraps Excel's .Sort method, with signature matching sSortedArray2
' -----------------------------------------------------------------------------------------------------------------------
Private Function SortWrap2(ArrayToSort, KeyCols, Ascendings, CaseSensitive As Boolean, NumHeaderRows As Long)

          Dim i As Long
          Dim NumKeyCols As Long
          Dim R As Range
          Dim Res
          Dim SUH As clsScreenUpdateHandler

1         Set SUH = CreateScreenUpdateHandler()

2         On Error GoTo ErrHandler
3         If TypeName(ArrayToSort) = "Range" Then ArrayToSort = ArrayToSort.Value2
4         NumKeyCols = sNRows(KeyCols)

5         With shEmptySheet
6             .Unprotect
7             .UsedRange.Clear
8             .UsedRange.EntireRow.Delete
9             Set R = .Cells(1, 1).Resize(sNRows(ArrayToSort), sNCols(ArrayToSort))
10            R.Value = sArrayExcelString(ArrayToSort)
11        End With

12        shEmptySheet.Sort.SortFields.Clear

13        For i = 1 To NumKeyCols
14            shEmptySheet.Sort.SortFields.Add Key:=R.Columns(KeyCols(i, 1)) _
                  , SortOn:=xlSortOnValues, order:=IIf(Ascendings(i, 1), xlAscending, xlDescending), DataOption:=xlSortNormal
15        Next i

16        With shEmptySheet.Sort
17            .SetRange R.Offset(NumHeaderRows).Resize(R.Rows.Count - NumHeaderRows)
18            .header = xlNo
19            .MatchCase = CaseSensitive
20            .Orientation = xlTopToBottom
21            .SortMethod = xlPinYin
22            .Apply
23        End With
24        SortWrap2 = R.Value
25        shEmptySheet.UsedRange.EntireRow.Delete
26        Res = shEmptySheet.UsedRange.Rows.Count

27        Exit Function
ErrHandler:
28        Throw "#SortWrap2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sEquals
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Returns TRUE if a is equal to b, FALSE otherwise. a and b may be numbers, strings,
'             Booleans or Excel error values, but not arrays. For testing equality of
'             arrays see ArrayEquals and sArraysIdentical.
'             Examples
'             sEquals(1,1) = TRUE
'             sEquals(#DIV0!,1) = FALSE
' Arguments
' a         : The first value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' b         : The second value to compare. Must be a single cell (not an array) but can contain numbers,
'             text, logical value or error values.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
'
'Note:        Avoids VBA booby trap that False = 0 and True = -1
' -----------------------------------------------------------------------------------------------------------------------
Function sEquals(a, b, Optional CaseSensitive As Boolean = False) As Variant
Attribute sEquals.VB_Description = "Returns TRUE if a is equal to b, FALSE otherwise. a and b may be numbers, strings, Booleans or Excel error values, but not arrays. For testing equality of arrays see ArrayEquals and sArraysIdentical.\nExamples\nsEquals(1,1) = TRUE\nsEquals(#DIV0!,1) = FALSE"
Attribute sEquals.VB_ProcData.VB_Invoke_Func = " \n27"
1         On Error GoTo ErrHandler
          Dim VTA As Long
          Dim VTB As Long

2         VTA = VarType(a)
3         VTB = VarType(b)
4         If VTA >= vbArray Or VTB >= vbArray Then
5             sEquals = "#sEquals: Function does not handle arrays. Use sArrayEquals or sArraysIdentical instead!"
6             Exit Function
7         End If

8         If VTA = VTB Then
9             If VTA = vbString And Not CaseSensitive Then
10                If Len(a) = Len(b) Then
11                    sEquals = UCase$(a) = UCase$(b)
12                Else
13                    sEquals = False
14                End If
15            Else
16                sEquals = (a = b)
17            End If
18        Else
19            If VTA = vbBoolean Or VTB = vbBoolean Or VTA = vbString Or VTB = vbString Then
20                sEquals = False
21            Else
22                sEquals = (a = b)
23            End If
24        End If
25        Exit Function
ErrHandler:
26        sEquals = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WrapSORT3
' Author     : Philip Swannell
' Date       : 11-Dec-2019
' Purpose    : Wraps SORTBY, but with signature matching that of sSortedArrayByAllCols
' -----------------------------------------------------------------------------------------------------------------------
Private Function WrapSORTBY2(ByVal ArrayToSort As Variant, Optional IgnoreFirstColumn As Boolean, Optional NumHeaderRows As Long = 0)
          Dim i As Long
          Dim MinCol As Long
          Dim NCols As Long
          Dim KeyCols As Variant
          Dim Ascending As Variant
          Const MAXCOLS = 10 'max number of columns supported by WrapSORTBY

1         On Error GoTo ErrHandler
2         MinCol = IIf(IgnoreFirstColumn, 2, 1)

3         For i = sNCols(ArrayToSort) To MinCol Step -MAXCOLS
4             NCols = i - MinCol + 1
5             If NCols > MAXCOLS Then NCols = MAXCOLS
6             KeyCols = sGrid(i - NCols + 1, CDbl(i), NCols)
7             Ascending = sReshape(True, NCols, 1)
8             ArrayToSort = WrapSORTBY(ArrayToSort, KeyCols, Ascending, NumHeaderRows)
9         Next

10        WrapSORTBY2 = ArrayToSort

11        Exit Function
ErrHandler:
12        WrapSORTBY2 = "#WrapSORTBY2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSortedArrayByAllCols
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : Sorts a multi-column array in ascending order. All columns (or all but the first column)
'             are used in order as keys for the sort.
' Arguments
' ArrayToSort: An array of arbitrary values.
' IgnoreFirstColumn: TRUE: The first column is used as the primary sort key. FALSE: The first column is ignored
'             as a sort key, and the second column is the primary sort key. This argument
'             is optional, defaulting to FALSE.
' CaseSensitive: Specifies if sorting of strings is case-sensitive.
' -----------------------------------------------------------------------------------------------------------------------
Public Function sSortedArrayByAllCols(ByVal ArrayToSort As Variant, Optional IgnoreFirstColumn As Boolean, Optional CaseSensitive As Boolean, Optional NumHeaderRows As Long = 0, Optional UseExcelSortMethod As Boolean)
Attribute sSortedArrayByAllCols.VB_Description = "Sorts a multi-column array in ascending order. All columns (or all but the first column) are used in order as keys for the sort."
Attribute sSortedArrayByAllCols.VB_ProcData.VB_Invoke_Func = " \n27"

          'Check the inputs
1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             sSortedArrayByAllCols = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         If UseExcelSortMethod Then
7             If ExcelSupportsSpill() Then
8                 If CaseSensitive Then ArrayToSort = EncodeArray(ArrayToSort)
9                 ArrayToSort = WrapSORTBY2(ArrayToSort, IgnoreFirstColumn, NumHeaderRows)
10                If CaseSensitive Then ArrayToSort = DecodeArray(ArrayToSort)
11                sSortedArrayByAllCols = ArrayToSort
12                Exit Function
13            ElseIf TypeName(Application.Caller) <> "Range" Then
                  Dim KeyCols, Ascendings, fcol As Double, lcol As Double
14                fcol = IIf(IgnoreFirstColumn, 2, 1)
15                lcol = sNCols(ArrayToSort)
16                KeyCols = sGrid(fcol, lcol, lcol - fcol + 1)
17                Ascendings = sReshape(True, lcol - fcol + 1, 1)
18                sSortedArrayByAllCols = SortWrap2(ArrayToSort, KeyCols, Ascendings, CaseSensitive, NumHeaderRows)
19                Exit Function
20            Else
21                If ExcelSupportsSpill() Then
22                    Throw "Cannot do case-sensitive sort using Excel sort method when function is called from spreadsheet"
23                Else
24                    Throw "Cannot sort using Excel sort method when function is called from spreadsheet"
25                End If
26            End If
27        End If

28        Force2DArrayR ArrayToSort

          'Copy to module-level variable, so that we can access from CompareRows3
29        m_IgnoreFirstColumn = IgnoreFirstColumn
30        m_ArrayToSort = ArrayToSort
31        m_NumCols = sNCols(m_ArrayToSort)

32        If m_NumCols = 1 Then
33            If m_IgnoreFirstColumn Then
34                sSortedArrayByAllCols = ArrayToSort
35                Exit Function
36            End If
37        End If

38        If NumHeaderRows < 0 Then
39            Throw "NumHeaderRows must be greater than or equal to zero"
40        ElseIf NumHeaderRows = 0 Then
41        ElseIf NumHeaderRows >= sNRows(ArrayToSort) Then
42            Throw "NumHeaderRows must be less than the number of rows in ArrayToSort"
43        End If

          Dim i As Long
          Dim NumCols As Long
          Dim NumRows As Long
          Dim vArray() As Long
44        NumRows = UBound(m_ArrayToSort, 1)
45        NumCols = UBound(m_ArrayToSort, 2)
46        ReDim vArray(1 To NumRows)
47        For i = 1 To NumRows
48            vArray(i) = i
49        Next

50        QuickSort3 vArray, 1 + NumHeaderRows, UBound(vArray), CaseSensitive

          'vArray now acts as an index for constructing ReturnArray
          Dim j As Long
          Dim ReturnArray() As Variant
51        ReDim ReturnArray(1 To UBound(m_ArrayToSort, 1), 1 To UBound(m_ArrayToSort, 2))
52        For i = 1 To NumRows
53            For j = 1 To NumCols
54                ReturnArray(i, j) = ArrayToSort(vArray(i), j)
55            Next j
56        Next i

57        sSortedArrayByAllCols = ReturnArray

58        Exit Function
ErrHandler:
59        sSortedArrayByAllCols = "#sSortedArrayByAllCols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSort3
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : sub-routine of sSortedArrayByAllCols
' -----------------------------------------------------------------------------------------------------------------------
Private Sub QuickSort3(vArray() As Long, inLow As Long, inHi As Long, CaseSensitive As Boolean)

          Dim pivot As Long
          Dim tmpHi As Long
          Dim tmpLow As Long
          Dim tmpSwap As Variant

1         On Error GoTo ErrHandler

2         tmpLow = inLow
3         tmpHi = inHi

4         pivot = vArray((inLow + inHi) \ 2)

5         Do While (tmpLow <= tmpHi)

6             Do While (CompareRows3(vArray(tmpLow), pivot, CaseSensitive) And tmpLow < inHi)
7                 tmpLow = tmpLow + 1
8             Loop

9             Do While (CompareRows3(pivot, vArray(tmpHi), CaseSensitive) And tmpHi > inLow)
10                tmpHi = tmpHi - 1
11            Loop

12            If (tmpLow <= tmpHi) Then
13                tmpSwap = vArray(tmpLow)
14                vArray(tmpLow) = vArray(tmpHi)
15                vArray(tmpHi) = tmpSwap
16                tmpLow = tmpLow + 1
17                tmpHi = tmpHi - 1
18            End If

19        Loop

20        If (inLow < tmpHi) Then QuickSort3 vArray, inLow, tmpHi, CaseSensitive
21        If (tmpLow < inHi) Then QuickSort3 vArray, tmpLow, inHi, CaseSensitive

22        Exit Sub
ErrHandler:
23        Throw "#QuickSort3 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompareRows3
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : Subroutine of QuickSort3 in turn a subroutine of sSortedArrayByAllCols
' -----------------------------------------------------------------------------------------------------------------------
Private Function CompareRows3(RowNo1 As Long, RowNo2 As Long, CaseSensitive As Boolean) As Boolean

1         On Error GoTo ErrHandler
          Dim i As Long

2         For i = IIf(m_IgnoreFirstColumn, 2, 1) To m_NumCols
3             If VariantLessThan(m_ArrayToSort(RowNo1, i), m_ArrayToSort(RowNo2, i), CaseSensitive) Then
4                 CompareRows3 = True
5                 Exit Function
6             ElseIf Not sEquals(m_ArrayToSort(RowNo1, i), m_ArrayToSort(RowNo2, i), CaseSensitive) Then
7                 CompareRows3 = False
8                 Exit Function
9             End If
10        Next i

11        CompareRows3 = RowNo1 < RowNo2        'Get a static sort this way

12        Exit Function
ErrHandler:
13        Throw "#CompareRows3 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSortMerge
' Author    : Philip Swannell
' Date      : 13-May-2015
' Purpose   : Collates data on a row-by-row basis. Rows are identified as similar if they have the same
'             entries in the key columns. The value columns are summed over rows with the
'             same keys. The output is a list of the (key, value) combinations, sorted by
'             key.
' Arguments
' TheArray  : Array of columns of data
' KeyColumns: List of columns to use as the "keys" (counting from 1). Can be a single integer for a
'             single column key, a column of integers for a set of columns, or string like
'             "2,5" for a set of columns. Elements of TheArray can be any type for key
'             columns.
' ValueColumns: List of columns to treat as the "values" of the input array's rows. Can be entered as
'             KeyColumns.
' Operations: The operation used where rows have the same key. Can be Sum, Count, Average, Max, Min,
'             SumOfNums, CountOfNums, MaxOfNums, MinOfNums - last 4 ignore non-numbers.
'             Pass a column array or comma-delimited string for different operations on
'             each value column.
' CaseSensitive: If TRUE, then key values which differ only in case are treated as different keys; if False
'             they are treated as being the same key.
'
' Notes     : Columns which are neither key columns nor value columns do not appear in the output,
'             otherwise the order of columns in the output is the same as the order of
'             column in the input. The output is sorted in ascending order on the key
'             columns.
'
'             Example
'             If TheArray is:
'             1    100
'             1    200
'             2    300
'             2    400
'             3    500
'             3    600
'
'             Then sSortMerge(TheArray,1,"2,2","Average,Max") yields:
'             1    150    200
'             2    350    400
'             3    550    600
' -----------------------------------------------------------------------------------------------------------------------
Function sSortMerge(TheArray As Variant, ByVal KeyColumns As Variant, ByVal ValueColumns As Variant, ByVal Operations As Variant, Optional CaseSensitive As Boolean)
Attribute sSortMerge.VB_Description = "Collates data on a row-by-row basis. Rows are identified as similar if they have the same entries in the key columns. The value columns are summed over rows with the same keys. The output is a list of the (key, value) combinations, sorted by key."
Attribute sSortMerge.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim c As Variant
          Dim D As Variant
          Dim DoingAverages As Boolean
          Dim i As Long
          Dim j As Long
          Dim KeyOutCols
          Dim NC As Long
          Dim NewLine As Boolean
          Dim NKC As Long
          Dim NR As Long
          Dim NVC As Long
          Dim OpCodes() As Long
          Dim ResultArray()
          Dim TempArray()
          Dim TheArraySorted
          Dim ValueOutCols
          Dim WriteLine As Long

          Const KeyColsError = "KeyColumns must be either a column array listing column numbers or a string such as ""1,2,3""."
          Const ValueColsError = "ValueColumns must be either a column array listing column numbers or a string such as ""4,5,6""."
          Const OperationsError = "Operations must be a column array or comma delimited string with allowed elements Sum, Count, Average, Median, Max, Min, SumOfNums, CountOfNums, AverageOfNums, MaxOfNums, MinOfNums"

1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             sSortMerge = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         If VarType(KeyColumns) = vbString Then
7             KeyColumns = sTokeniseString(CStr(KeyColumns))
8             For i = 1 To sNRows(KeyColumns)
9                 If Not IsNumeric(KeyColumns(i, 1)) Then Throw KeyColsError
10                KeyColumns(i, 1) = CDbl(KeyColumns(i, 1))
11            Next i
12        End If

13        If VarType(ValueColumns) = vbString Then
14            ValueColumns = sTokeniseString(CStr(ValueColumns))
15            For i = 1 To sNRows(ValueColumns)
16                If Not IsNumeric(ValueColumns(i, 1)) Then Throw ValueColsError
17                ValueColumns(i, 1) = CDbl(ValueColumns(i, 1))
18            Next i
19        End If
20        If VarType(Operations) = vbString Then
21            Operations = sTokeniseString(CStr(Operations))
22        End If

23        Force2DArrayRMulti TheArray, KeyColumns, ValueColumns, Operations
24        NR = sNRows(TheArray)
25        NC = sNCols(TheArray)
26        If sNCols(KeyColumns) <> 1 Then Throw KeyColsError
27        If sNCols(ValueColumns) <> 1 Then Throw ValueColsError
28        If sNCols(Operations) <> 1 Then Throw OperationsError
29        If sNRows(sRemoveDuplicates(KeyColumns)) <> sNRows(KeyColumns) Then Throw "KeyColumns cannot contain repeated elements"
30        If sNRows(Operations) = 1 And sNRows(ValueColumns) > 1 Then
31            Operations = sReshape(Operations, sNRows(ValueColumns), 1)
32        ElseIf sNRows(Operations) <> sNRows(ValueColumns) Then
33            Throw "There must be the same number of Operations as there are ValueColumns"
34        End If

35        For Each c In KeyColumns
36            If IsNumeric(c) Then c = CDbl(c)
37            If Not IsNumber(c) Then Throw KeyColsError
38            If c < 1 Or c > NC Then Throw KeyColsError
39            If c <> CInt(c) Then Throw KeyColsError
40        Next c
41        For Each c In ValueColumns
42            If IsNumeric(c) Then c = CDbl(c)
43            If Not IsNumber(c) Then Throw ValueColsError
44            If c < 1 Or c > NC Then Throw ValueColsError
45            If c <> CInt(c) Then Throw ValueColsError
46        Next c
47        For Each c In KeyColumns
48            For Each D In ValueColumns
49                If c = D Then Throw "KeyColumns and ValueColumns must be disjoint but " + CStr(c) + " appears in both of them"
50            Next D
51        Next c
52        NKC = sNRows(KeyColumns): NVC = sNRows(ValueColumns)

53        ReDim OpCodes(1 To NVC)

54        For i = 1 To NVC
55            Select Case LCase$(Operations(i, 1))
                  Case "sum", "add"
56                    OpCodes(i) = 0
57                Case "count"
58                    OpCodes(i) = 1
59                Case "max"
60                    OpCodes(i) = 2
61                Case "min"
62                    OpCodes(i) = 3
63                Case "sumofnums"
64                    OpCodes(i) = 4
65                Case "maxofnums"
66                    OpCodes(i) = 5
67                Case "minofnums"
68                    OpCodes(i) = 6
69                Case "average", "mean"
70                    OpCodes(i) = 7
71                    DoingAverages = True
72                Case "averageofnums", "meanofnums"
73                    OpCodes(i) = 8
74                    DoingAverages = True
75                Case "countofnums"
76                    OpCodes(i) = 9
77                Case "median"
78                    OpCodes(i) = 10
79                Case Else
80                    Throw OperationsError
81            End Select
82        Next i

83        For j = 1 To NVC
84            Select Case OpCodes(j)
                  Case 0, 2, 3, 7, 10
85                    For i = 1 To NR
86                        Select Case VarType(TheArray(i, ValueColumns(j, 1)))
                              Case vbString, vbError, vbBoolean
87                                Throw "Value columns must contain numbers only. Non number found at line " + CStr(i) + ", column " + CStr(ValueColumns(j, 1)) + " Hint: Try ""tolerant"" operations such as SumOfNums"
88                        End Select
89                    Next i
90            End Select
91        Next j
          'FINISHED INPUT CHECKING

          'Construct the arrays to index into the columns for writing to...
92        ReDim TempArray(1 To NKC + NVC, 1 To 3)
93        For i = 1 To NKC
94            TempArray(i, 1) = KeyColumns(i, 1)
95            TempArray(i, 2) = True
96        Next i
97        For i = 1 To NVC
98            TempArray(NKC + i, 1) = ValueColumns(i, 1)
99            TempArray(NKC + i, 2) = False
100       Next i
101       TempArray = sSortedArray(TempArray)
102       For i = 1 To NKC + NVC
103           TempArray(i, 3) = i
104       Next i
105       KeyOutCols = sMChoose(sSubArray(TempArray, 1, 3, , 1), sSubArray(TempArray, 1, 2, , 1))
106       ValueOutCols = sMChoose(sSubArray(TempArray, 1, 3, , 1), sArrayNot(sSubArray(TempArray, 1, 2, , 1)))

107       TheArraySorted = sSortedArray2(TheArray, KeyColumns, sReshape(True, sNRows(KeyColumns), 1), CaseSensitive)

108       ReDim ResultArray(1 To NR, 1 To NKC + NVC)
109       WriteLine = 1
110       For j = 1 To NKC
111           ResultArray(WriteLine, KeyOutCols(j, 1)) = TheArraySorted(1, KeyColumns(j, 1))
112       Next j
113       For j = 1 To NVC
114           ResultArray(WriteLine, ValueOutCols(j, 1)) = InitialiseElement(TheArraySorted(1, ValueColumns(j, 1)), OpCodes(j))
115       Next j

116       For i = 2 To NR
117           NewLine = False
118           For j = NKC To 1 Step -1
119               If Not sEquals(TheArraySorted(i, KeyColumns(j, 1)), TheArraySorted(i - 1, KeyColumns(j, 1)), CaseSensitive) Then
120                   WriteLine = WriteLine + 1
121                   NewLine = True
122                   Exit For
123               End If
124           Next j
125           If NewLine Then
126               For j = 1 To NKC
127                   ResultArray(WriteLine, KeyOutCols(j, 1)) = TheArraySorted(i, KeyColumns(j, 1))
128               Next j
129               For j = 1 To NVC
130                   ResultArray(WriteLine, ValueOutCols(j, 1)) = InitialiseElement(TheArraySorted(i, ValueColumns(j, 1)), OpCodes(j))
131               Next j
132           Else
133               For j = 1 To NVC
134                   Select Case OpCodes(j)
                          Case 0        'Sum
135                           ResultArray(WriteLine, ValueOutCols(j, 1)) = TheArraySorted(i, ValueColumns(j, 1)) + ResultArray(WriteLine, ValueOutCols(j, 1))
136                       Case 1        'Count
137                           ResultArray(WriteLine, ValueOutCols(j, 1)) = 1 + ResultArray(WriteLine, ValueOutCols(j, 1))
138                       Case 2        'Max
139                           ResultArray(WriteLine, ValueOutCols(j, 1)) = SafeMax(TheArraySorted(i, ValueColumns(j, 1)), ResultArray(WriteLine, ValueOutCols(j, 1)))
140                       Case 3        'Min
141                           ResultArray(WriteLine, ValueOutCols(j, 1)) = SafeMin(TheArraySorted(i, ValueColumns(j, 1)), ResultArray(WriteLine, ValueOutCols(j, 1)))
142                       Case 4        'SumOfNums
143                           ResultArray(WriteLine, ValueOutCols(j, 1)) = SumON(TheArraySorted(i, ValueColumns(j, 1)), ResultArray(WriteLine, ValueOutCols(j, 1)))
144                       Case 5        'MaxOfNums
145                           ResultArray(WriteLine, ValueOutCols(j, 1)) = MaxON(TheArraySorted(i, ValueColumns(j, 1)), ResultArray(WriteLine, ValueOutCols(j, 1)))
146                       Case 6        'MinOfNums
147                           ResultArray(WriteLine, ValueOutCols(j, 1)) = MinON(TheArraySorted(i, ValueColumns(j, 1)), ResultArray(WriteLine, ValueOutCols(j, 1)))
148                       Case 9        'CountOfNums
149                           If IsNumberOrDate(TheArraySorted(i, ValueColumns(j, 1))) Then
150                               ResultArray(WriteLine, ValueOutCols(j, 1)) = 1 + ResultArray(WriteLine, ValueOutCols(j, 1))
151                           End If
152                       Case 7        'Average
153                           ResultArray(WriteLine, ValueOutCols(j, 1)) = AddToPair(CDbl(TheArraySorted(i, ValueColumns(j, 1))), ResultArray(WriteLine, ValueOutCols(j, 1)))
154                       Case 8        'AverageOfNums
155                           If IsNumberOrDate(TheArraySorted(i, ValueColumns(j, 1))) Then
156                               ResultArray(WriteLine, ValueOutCols(j, 1)) = AddToPair(CDbl(TheArraySorted(i, ValueColumns(j, 1))), ResultArray(WriteLine, ValueOutCols(j, 1)))
157                           End If
158                   End Select
159               Next j
160           End If
161       Next i
162       If DoingAverages Then
163           For j = 1 To NVC
164               Select Case OpCodes(j)
                      Case 7, 8        'Average,AverageOfNums
165                       For i = 1 To WriteLine
166                           ResultArray(i, ValueOutCols(j, 1)) = CollapsePair(ResultArray(i, ValueOutCols(j, 1)))
167                       Next i
168               End Select
169           Next j
170       End If

171       sSortMerge = sSubArray(ResultArray, 1, 1, WriteLine)
172       Exit Function
ErrHandler:
173       sSortMerge = "#sSortMerge (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CollapsePair
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : For averages we use pairs (two element arrays) to keep a running total and
'             a running count so the pair needs to be "collapsed" back to the required average...
' -----------------------------------------------------------------------------------------------------------------------
Private Function CollapsePair(Pair As Variant)
1         On Error GoTo ErrHandler
2         If Pair(1) = 0 Then
3             CollapsePair = "#Nothing to average!"
4         Else
5             CollapsePair = Pair(0) / Pair(1)
6         End If
7         Exit Function
ErrHandler:
8         Throw "#CollapsePair (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddToPair
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : To handle averages, we use pairs (two elemnt arrays) as the thing that gets accumulated...
' -----------------------------------------------------------------------------------------------------------------------
Private Function AddToPair(AddThis As Double, ByVal Pair As Variant)
1         On Error GoTo ErrHandler
2         Pair(0) = Pair(0) + AddThis
3         Pair(1) = Pair(1) + 1
4         AddToPair = Pair
5         Exit Function
ErrHandler:
6         Throw "#AddToPair (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InitialiseElement
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : Sub of SSortMerge - when we encounter the first element to "accumulate", what
'             should the "sum/max/min/count..." be set to? Return may be a number of two-element array (for OpCode is 7 or 8)
' -----------------------------------------------------------------------------------------------------------------------
Private Function InitialiseElement(InElement, OpCode As Long)

1         On Error GoTo ErrHandler
2         Select Case OpCode
              Case 1        'count
3                 InitialiseElement = 1
4             Case 0, 2, 3        'sum, max, min
5                 InitialiseElement = InElement
6             Case 7        'average - populate with two element array giving running total and count
7                 InitialiseElement = VBA.Array(InElement, 1)
8             Case Else
9                 If IsNumberOrDate(InElement) Then
10                    Select Case OpCode
                          Case 9        ' countofnums
11                            InitialiseElement = 1
12                        Case 4, 5, 6        'sumofnums, maxofnums, minofnums
13                            InitialiseElement = InElement
14                        Case 8        'averageofnums - populate with two element array giving running total and count
15                            InitialiseElement = VBA.Array(InElement, 1)
16                    End Select
17                Else
18                    Select Case OpCode
                          Case 9, 4        ' countofnums, sumofnums
19                            InitialiseElement = 0
20                        Case 5, 6        ' maxofnums, minofnums
21                            InitialiseElement = "#No numbers found!"
22                        Case 8        'averageofnums - populate with two element array giving running total and count
23                            InitialiseElement = VBA.Array(0, 0)
24                    End Select
25                End If
26        End Select
27        Exit Function
ErrHandler:
28        Throw "#InitialiseElement (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SumON
' Author    : Philip Swannell
' Date      : 11-Jan-2016
' Purpose   : Sub routine of sSortMerge. Not symmetric in A & B. A is the value being added, B is the running total
' -----------------------------------------------------------------------------------------------------------------------
Private Function SumON(a, b)
1         If IsNumberOrDate(a) Then
2             If IsNumberOrDate(b) Then
3                 SumON = a + b
4             Else
5                 SumON = a
6             End If
7         Else
8             SumON = b
9         End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MaxON
' Author    : Philip Swannell
' Date      : 11-Jan-2016
' Purpose   : Sub routine of sSortMerge. Not symmetric in A & B. A is the value being added, B is the running maximum
' -----------------------------------------------------------------------------------------------------------------------
Private Function MaxON(a, b)
1         If IsNumberOrDate(a) Then
2             If IsNumberOrDate(b) Then
3                 If a > b Then
4                     MaxON = a
5                 Else
6                     MaxON = b
7                 End If
8             Else
9                 MaxON = a
10            End If
11        Else
12            MaxON = b
13        End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MinON
' Author    : Philip Swannell
' Date      : 11-Jan-2016
' Purpose   : Sub routine of sSortMerge. Not symmetric in A & B. A is the value being added, B is the running minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function MinON(a, b)
1         If IsNumberOrDate(a) Then
2             If IsNumberOrDate(b) Then
3                 If a > b Then
4                     MinON = b
5                 Else
6                     MinON = a
7                 End If
8             Else
9                 MinON = a
10            End If
11        Else
12            MinON = b
13        End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSortedArray2
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Alternative version of sSortedArray, allowing sorting on more than three key columns.
' Arguments
' ArrayToSort: An array of arbitrary values.
' KeyCols   : Identifies which columns of ArrayToSort are to be used as keys in the sort operation,
'             counting from 1 as the left-most column. Should be entered either as a column
'             array of integers or as a comma delimited string such as "1,3,5".
' Ascendings: Specifies whether the sort should be Ascending (TRUE) or Descending (FALSE) on the
'             corresponding KeyColumn. Should be specified either as a column array of
'             logical values or as a comma-delimited string such as "TRUE,FALSE,TRUE"
' CaseSensitive: Specifies if sorting of strings is case-sensitive.
' -----------------------------------------------------------------------------------------------------------------------
Function sSortedArray2(ArrayToSort, KeyCols, Ascendings, CaseSensitive As Boolean, Optional NumHeaderRows As Long = 0, Optional UseExcelSortMethod As Boolean = False)
Attribute sSortedArray2.VB_Description = "Alternative version of sSortedArray, allowing sorting on more than three key columns."
Attribute sSortedArray2.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim vArray() As Long

          Const KeyColsError = "KeyCols must be either a column array listing column numbers which which are to be the key columns for sorting a string such as ""1,2,3"". "
          Const AscendingsError = "Ascendings must be a column array of logicals. Each element determines if the corresponding key column is sorted ascending (TRUE) or descending (FALSE). Strings such as ""T,T,T"" are also allowed."

1         If FunctionWizardActive() Then
2             sSortedArray2 = "#Disabled in Function Dialog!"
3             Exit Function
4         End If

5         On Error GoTo ErrHandler
6         If VarType(KeyCols) = vbString Then
7             KeyCols = sTokeniseString(CStr(KeyCols))
8             For i = 1 To sNRows(KeyCols)
9                 If Not IsNumeric(KeyCols(i, 1)) Then Throw KeyColsError
10                KeyCols(i, 1) = CDbl(KeyCols(i, 1))
11            Next i
12        End If

13        If VarType(Ascendings) = vbString Then
14            Ascendings = sTokeniseString(CStr(Ascendings))
15            For i = 1 To sNRows(Ascendings)
16                If UCase$(Left$(Ascendings(i, 1), 1)) = "T" Then
17                    Ascendings(i, 1) = True
18                ElseIf UCase$(Left$(Ascendings(i, 1), 1)) = "F" Then
19                    Ascendings(i, 1) = False
20                Else
21                    Throw AscendingsError
22                    Ascendings(i, 1) = CDbl(Ascendings(i, 1))
23                End If
24            Next i
25        End If

26        Force2DArrayRMulti ArrayToSort, KeyCols, Ascendings
27        NC = sNCols(ArrayToSort): NR = sNRows(ArrayToSort)

28        If sNCols(KeyCols) <> 1 Then Throw KeyColsError
29        If sNCols(Ascendings) <> 1 Then Throw AscendingsError
30        If sNRows(KeyCols) <> sNRows(Ascendings) Then Throw "No of elements in KeyCols must equal number of elements in Ascendings"
31        For i = 1 To sNRows(Ascendings)
32            If VarType(Ascendings(i, 1)) <> vbBoolean Then Throw AscendingsError
33        Next i
34        For i = 1 To sNRows(KeyCols)
35            If Not IsNumber(KeyCols(i, 1)) Then Throw AscendingsError
36            If KeyCols(i, 1) < 1 Then Throw "KeyCols cannot be less than 1"
37            If KeyCols(i, 1) > NC Then Throw "KeyCols cannot exceed the number of columns in ArrayToSort"
38            If KeyCols(i, 1) <> CInt(KeyCols(i, 1)) Then Throw "KeyCols must contain only integers from 1 to " & CStr(NC)
39        Next i
40        If sNRows(sRemoveDuplicates(KeyCols)) <> sNRows(KeyCols) Then Throw "KeyCols cannot contain repeated elements"

41        If NumHeaderRows < 0 Then
42            Throw "NumHeaderRows must be greater than or equal to zero"
43        ElseIf NumHeaderRows = 0 Then
44        ElseIf NumHeaderRows >= sNRows(ArrayToSort) Then
45            Throw "NumHeaderRows must be less than the number of rows in ArrayToSort"
46        End If

          'Use native Excel sorting if we can...
47        If UseExcelSortMethod Then
48            If ExcelSupportsSpill() Then
49                If CaseSensitive Then ArrayToSort = EncodeArray(ArrayToSort)
50                ArrayToSort = WrapSORTBY(ArrayToSort, KeyCols, Ascendings, NumHeaderRows)
51                If CaseSensitive Then ArrayToSort = DecodeArray(ArrayToSort)
52                sSortedArray2 = ArrayToSort
53                Exit Function
54            ElseIf TypeName(Application.Caller) <> "Range" Then
55                sSortedArray2 = SortWrap2(ArrayToSort, KeyCols, Ascendings, CaseSensitive, NumHeaderRows)
56                Exit Function
57            Else
58                If ExcelSupportsSpill() Then
59                    Throw "Cannot do case-sensitive sort using Excel sort method when function is called from spreadsheet"
60                Else
61                    Throw "Cannot sort using Excel sort method when function is called from spreadsheet"
62                End If
63            End If
64        End If

65        m_ArrayToSort = ArrayToSort
66        mKeyCols = KeyCols
67        mAscendings = Ascendings
68        m_NumCols = sNRows(mKeyCols)

69        ReDim vArray(1 To NR)
70        For i = 1 To NR
71            vArray(i) = i
72        Next

73        QuickSort4 vArray, NumHeaderRows + 1, UBound(vArray), CaseSensitive

          Dim j As Long
          Dim ReturnArray() As Variant

74        ReDim ReturnArray(1 To NR, 1 To NC)
75        For i = 1 To NR
76            For j = 1 To NC
77                ReturnArray(i, j) = ArrayToSort(vArray(i), j)
78            Next j
79        Next i
80        sSortedArray2 = ReturnArray

81        Exit Function
ErrHandler:
82        sSortedArray2 = "#sSortedArray2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSort4
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : sub-routine of sSortedArray2
' -----------------------------------------------------------------------------------------------------------------------
Private Sub QuickSort4(vArray() As Long, inLow As Long, inHi As Long, CaseSensitive As Boolean)

          Dim pivot As Long
          Dim tmpHi As Long
          Dim tmpLow As Long
          Dim tmpSwap As Variant

1         On Error GoTo ErrHandler

2         tmpLow = inLow
3         tmpHi = inHi

4         pivot = vArray((inLow + inHi) \ 2)

5         Do While (tmpLow <= tmpHi)

6             Do While (CompareRows4(vArray(tmpLow), pivot, CaseSensitive) And tmpLow < inHi)
7                 tmpLow = tmpLow + 1
8             Loop

9             Do While (CompareRows4(pivot, vArray(tmpHi), CaseSensitive) And tmpHi > inLow)
10                tmpHi = tmpHi - 1
11            Loop

12            If (tmpLow <= tmpHi) Then
13                tmpSwap = vArray(tmpLow)
14                vArray(tmpLow) = vArray(tmpHi)
15                vArray(tmpHi) = tmpSwap
16                tmpLow = tmpLow + 1
17                tmpHi = tmpHi - 1
18            End If

19        Loop

20        If (inLow < tmpHi) Then QuickSort4 vArray, inLow, tmpHi, CaseSensitive
21        If (tmpLow < inHi) Then QuickSort4 vArray, tmpLow, inHi, CaseSensitive

22        Exit Sub
ErrHandler:
23        Throw "#QuickSort4 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompareRows4
' Author    : Philip Swannell
' Date      : 18-Apr-2015
' Purpose   : Subroutine of QuickSort4 in turn a subroutine of sSortedArray2
' -----------------------------------------------------------------------------------------------------------------------
Private Function CompareRows4(RowNo1 As Long, RowNo2 As Long, CaseSensitive As Boolean) As Boolean

1         On Error GoTo ErrHandler
          Dim i As Long

2         For i = 1 To m_NumCols
3             If VariantLessThan(m_ArrayToSort(RowNo1, mKeyCols(i, 1)), m_ArrayToSort(RowNo2, mKeyCols(i, 1)), CaseSensitive) Then
4                 CompareRows4 = mAscendings(i, 1)
5                 Exit Function
6             ElseIf Not sEquals(m_ArrayToSort(RowNo1, mKeyCols(i, 1)), m_ArrayToSort(RowNo2, mKeyCols(i, 1)), CaseSensitive) Then
7                 CompareRows4 = Not (mAscendings(i, 1))
8                 Exit Function
9             End If
10        Next i

11        CompareRows4 = RowNo1 < RowNo2        'Get a static sort this way

12        Exit Function
ErrHandler:
13        Throw "#CompareRows4 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedures : myEncode, myDecode, EncodeArray, DecodeArray
' Author     : Philip Swannell
' Date       : 11-Dec-2019
' Purpose    : Encode a string in such a way that sorting of the encoded strings is equivalent to case-sensitve sorting of
'              the unencoded strings. i.e. encode > sort > unencode is equivalent to case-sensitive sort.
'Drawback: Only ascii strings are coped with, not wide-character strings...
' -----------------------------------------------------------------------------------------------------------------------
Private Function myEncode(x As String)
          Dim y As String, i As Long, N As Long
          Static Table
1         On Error GoTo ErrHandler
2         N = Len(x)
3         y = String(3 * N, " ")

4         If IsEmpty(Table) Then
5             Table = VBA.Array( _
                  "000", "002", "003", "004", "005", "006", "007", "008", "009", "035", "036", "037", "038", "039", "010", "011", "012", "013", "014", _
                  "015", "016", "017", "018", "019", "020", "021", "022", "023", "024", "025", "026", "027", "033", "040", "041", "042", "043", _
                  "044", "045", "029", "046", "047", "048", "088", "049", "030", "050", "051", "115", "119", "121", "123", "125", "126", "127", _
                  "128", "129", "130", "052", "053", "089", "090", "091", "054", "055", "132", "149", "151", "155", "159", "169", "172", "174", _
                  "176", "186", "188", "190", "192", "194", "198", "215", "217", "219", "221", "226", "231", "241", "243", "245", "247", "253", _
                  "056", "057", "058", "059", "061", "062", "131", "148", "150", "154", "158", "168", "171", "173", "175", "185", "187", "189", _
                  "191", "193", "197", "214", "216", "218", "220", "225", "230", "240", "242", "244", "246", "252", "063", "064", "065", "066", _
                  "028", "087", "110", "077", "170", "080", "105", "106", "107", "060", "109", "223", "081", "213", "111", "255", "112", "113", _
                  "075", "076", "078", "079", "108", "031", "032", "074", "229", "222", "082", "212", "114", "254", "251", "034", "067", "083", _
                  "084", "085", "086", "068", "097", "069", "098", "133", "093", "099", "001", "100", "070", "101", "092", "122", "124", "071", _
                  "102", "103", "104", "072", "120", "199", "094", "116", "117", "118", "073", "137", "135", "139", "143", "141", "145", "147", _
                  "153", "163", "161", "165", "167", "180", "178", "182", "184", "157", "196", "203", "201", "205", "209", "207", "095", "211", _
                  "235", "233", "237", "239", "249", "228", "224", "136", "134", "138", "142", "140", "144", "146", "152", "162", "160", "164", _
                  "166", "179", "177", "181", "183", "156", "195", "202", "200", "204", "208", "206", "096", "210", "234", "232", "236", "238", _
                  "248", "227", "250")
6         End If

7         For i = 1 To N
8             Mid$(y, 3 * i - 2, 3) = Table(Asc(Mid(x, i, 1)))
9         Next
10        myEncode = y
11        Exit Function
ErrHandler:
12        Throw "#myEncode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function myDecode(x As String)
          Static Table
          Dim N As Long
          Dim y As String
          Dim i As Long

1         On Error GoTo ErrHandler
2         If IsEmpty(Table) Then
3             Table = VBA.Array(0, 173, 1, 2, 3, 4, 5, 6, 7, 8, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 39, 45, 150, 151, 32, 160, 9, 10, 11, 12, 13, 33, 34, 35, _
                  36, 37, 38, 40, 41, 42, 44, 46, 47, 58, 59, 63, 64, 91, 92, 93, 94, 136, 95, 96, 123, 124, 125, 126, 161, 166, 168, 175, 180, 184, 191, 152, 145, 146, 130, 147, 148, 132, 139, 155, _
                  162, 163, 164, 165, 128, 43, 60, 61, 62, 177, 171, 187, 215, 247, 167, 169, 172, 174, 176, 181, 182, 183, 133, 134, 135, 149, 137, 129, 141, 143, 144, 157, 48, 188, 189, 190, 49, _
                  185, 50, 178, 51, 179, 52, 53, 54, 55, 56, 57, 97, 65, 170, 225, 193, 224, 192, 226, 194, 228, 196, 227, 195, 229, 197, 230, 198, 98, 66, 99, 67, 231, 199, 100, 68, 240, 208, 101, 69, _
                  233, 201, 232, 200, 234, 202, 235, 203, 102, 70, 131, 103, 71, 104, 72, 105, 73, 237, 205, 236, 204, 238, 206, 239, 207, 106, 74, 107, 75, 108, 76, 109, 77, 110, 78, 241, 209, 111, 79, _
                  186, 243, 211, 242, 210, 244, 212, 246, 214, 245, 213, 248, 216, 156, 140, 112, 80, 113, 81, 114, 82, 115, 83, 154, 138, 223, 116, 84, 254, 222, 153, 117, 85, 250, 218, 249, 217, 251, 219, _
                  252, 220, 118, 86, 119, 87, 120, 88, 121, 89, 253, 221, 255, 159, 122, 90, 158, 142)
4         End If

5         N = Len(x) / 3
6         y = String(N, " ")

7         For i = 1 To N
8             Mid$(y, i, 1) = Chr$(Table(CInt(Mid(x, i * 3 - 2, 3))))
9         Next i
10        myDecode = y
11        Exit Function
ErrHandler:
12        Throw "#myDecode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function EncodeArray(ByVal x)
          Dim NR As Long, NC As Long, i As Long, j As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR x, NR, NC

3         For i = 1 To NR
4             For j = 1 To NC
5                 If VarType(x(i, j)) = vbString Then
6                     x(i, j) = myEncode(CStr(x(i, j)))
7                 End If
8             Next
9         Next
10        EncodeArray = x
11        Exit Function
ErrHandler:
12        Throw "#EncodeArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function DecodeArray(ByVal x)
          Dim NR As Long, NC As Long, i As Long, j As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR x, NR, NC

3         For i = 1 To NR
4             For j = 1 To NC
5                 If VarType(x(i, j)) = vbString Then
6                     x(i, j) = myDecode(CStr(x(i, j)))
7                 End If
8             Next
9         Next

10        DecodeArray = x

11        Exit Function
ErrHandler:
12        Throw "#DecodeArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub testencodedecode()
          Dim i As Long
1         For i = 1 To 255
2             If myDecode(myEncode(Chr(i))) <> Chr(i) Then Stop
3         Next
End Sub
