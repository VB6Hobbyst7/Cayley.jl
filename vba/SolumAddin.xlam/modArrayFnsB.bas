Attribute VB_Name = "modArrayFnsB"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modArrayFunctionsB
' Author    : Philip Swannell
' Date      : 24-May-2015
' Purpose   : More array functions
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMChoose
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : The input ChooseVector must be a column array of logicals with the same height as
'             TheArray. The function returns those rows of TheArray for which the
'             corresponding flag in ChooseVector is TRUE. If no flags are TRUE, the
'             function returns an error.
' Arguments
' TheArray  : An array of arbitrary values.
' ChooseVector: A column array of TRUE and FALSE with the same height as TheArray.
' -----------------------------------------------------------------------------------------------------------------------
Function sMChoose(ByVal TheArray As Variant, ByVal ChooseVector As Variant)
Attribute sMChoose.VB_Description = "The input ChooseVector must be a column array of logicals with the same height as TheArray. The function returns those rows of TheArray for which the corresponding flag in ChooseVector is TRUE. If no flags are TRUE, the function returns an error."
Attribute sMChoose.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NROut As Long
          Dim OutputArray As Variant
          Dim WriteRow As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray, NR, NC
3         Force2DArrayR ChooseVector

          'Test Inputs
4         If UBound(ChooseVector, 1) <> NR Then
5             sMChoose = "#ChooseVector and TheArray must have the same number of rows!"
6             Exit Function
7         ElseIf UBound(ChooseVector, 2) <> 1 Then
8             sMChoose = "#ChooseVector must have one column only!"
9             Exit Function
10        End If

11        For i = 1 To UBound(ChooseVector, 1)
12            If VarType(ChooseVector(i, 1)) = vbBoolean Then
13                If ChooseVector(i, 1) Then
14                    NROut = NROut + 1
15                End If
16            Else
17                sMChoose = "#ChooseVector must contain only True or False!"
18                Exit Function
19            End If
20        Next i

21        If NROut = 0 Then
22            sMChoose = "#Nothing to include!"
23            Exit Function
24        ElseIf NROut = NR Then
25            sMChoose = TheArray
26            Exit Function
27        End If

28        ReDim OutputArray(1 To NROut, 1 To NC)

29        For i = 1 To UBound(ChooseVector, 1)
30            If ChooseVector(i, 1) Then
31                WriteRow = WriteRow + 1
32                For j = 1 To NC
33                    OutputArray(WriteRow, j) = TheArray(i, j)
34                Next j
35            End If
36        Next i

37        sMChoose = OutputArray

38        Exit Function
ErrHandler:
39        sMChoose = "#sMChoose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNCols
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Function sNCols(Optional TheArray) As Long
Attribute sNCols.VB_Description = "Returns the number of columns in TheArray."
Attribute sNCols.VB_ProcData.VB_Invoke_Func = " \n24"
1         If TypeName(TheArray) = "Range" Then
2             sNCols = TheArray.Columns.Count
3         ElseIf IsMissing(TheArray) Then
4             sNCols = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNCols = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
10                Case Else
11                    sNCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNRows
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
' -----------------------------------------------------------------------------------------------------------------------
Function sNRows(Optional TheArray) As Long
Attribute sNRows.VB_Description = "Returns the number of rows in TheArray."
Attribute sNRows.VB_ProcData.VB_Invoke_Func = " \n24"
1         If TypeName(TheArray) = "Range" Then
2             sNRows = TheArray.Rows.Count
3         ElseIf IsMissing(TheArray) Then
4             sNRows = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             sNRows = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     sNRows = 1
10                Case Else
11                    sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
12            End Select
13        End If
End Function

Function sRandomStrings(NumStrings As Long, StringLength As Long, ByVal MinChar As String, ByVal MaxChar As String, Optional Seed As Variant)
Attribute sRandomStrings.VB_Description = "Returns a colum of random strings. There are NumStrings returned, each of length StringLength and each character chosen at random from the ascii character set between MinChar and MaxChar."
Attribute sRandomStrings.VB_ProcData.VB_Invoke_Func = " \n28"

          Dim i As Long
          Dim j As Long
          Dim Randoms As Variant
          Dim Result
          Dim x As Double
          Dim y As Double

1         On Error GoTo ErrHandler

2         If NumStrings <= 0 Then Throw "NumStrings must be positive"
3         If StringLength <= 0 Then Throw "StringLength must be positive"
4         If Len(MinChar) <> 1 Then Throw "MinChar must be a single character"
5         If Len(MaxChar) <> 1 Then Throw "MaxChar must be a single character"
6         If Asc(MinChar) > Asc(MaxChar) Then
              Dim tmp As String
7             tmp = MinChar
8             MinChar = MaxChar
9             MaxChar = tmp
10        End If

11        x = Asc(MinChar) - 0.5
12        y = Asc(MaxChar) - Asc(MinChar) + 1

13        Randoms = ThrowIfError(sRandomVariable(NumStrings, StringLength, "Uniform", , , Seed))
14        Result = sReshape(String(StringLength, " "), NumStrings, 1)

15        For i = 1 To NumStrings
16            For j = 1 To StringLength
17                Mid$(Result(i, 1), j, 1) = Chr$(x + y * Randoms(i, j))
18            Next j
19        Next i

20        sRandomStrings = Result

21        Exit Function
ErrHandler:
22        Throw "#sRandomStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRemoveDuplicates
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Returns an array listing the unique elements in the input single-column array. Optionally
'             can sort the return as well.
' Arguments
' Column    : The column from which to remove duplicated elements.
' SortAsWell: TRUE if the array is to be returned in sorted order. If omitted defaults to FALSE. Note
'             that the function executes faster if this argument is passed as TRUE.
' CaseSensitive: TRUE if comparison of strings to determine duplication is to be case-sensitive. If omitted
'             defaults to FALSE for case insensitive comparison of strings.
' -----------------------------------------------------------------------------------------------------------------------
Function sRemoveDuplicates(ByVal Column As Variant, _
          Optional SortAsWell As Boolean, _
          Optional CaseSensitive As Boolean = False)
Attribute sRemoveDuplicates.VB_Description = "Returns an array listing the unique elements in the input single-column array. Optionally can sort the return as well."
Attribute sRemoveDuplicates.VB_ProcData.VB_Invoke_Func = " \n27"

          Dim ChooseVector As Variant
          Dim ColumnNumberForFirstSort As Long
          Dim i As Long
          Dim N As Long
          Dim TempArray() As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR Column
3         N = sNRows(Column)
4         If sNCols(Column) <> 1 Then
5             sRemoveDuplicates = "#Column must be an array of data with only one column!"
6             Exit Function
7         End If

8         If CaseSensitive = False Then
9             If ExcelSupportsSpill() Then
10                sRemoveDuplicates = Application.WorksheetFunction.Unique(Column)
11                If SortAsWell Then
12                    sRemoveDuplicates = Application.WorksheetFunction.Sort(sRemoveDuplicates)
13                End If
14                Exit Function
15            Else
                  'In certain circumstances we can use the RemoveDuplicates method of the Range object. _
                   This is faster for arrays with more than approx 200 rows and for 100,000 rows is about _
                   13 times faster
16                If sNRows(Column) > 200 Then
17                    If TypeName(Application.Caller) <> "Range" Then
18                        sRemoveDuplicates = RemoveDuplicatesXL(Column, SortAsWell)
19                        Exit Function
20                    End If
21                End If
22            End If
23        End If

24        If Not CaseSensitive Then
25            ColumnNumberForFirstSort = 2
26            Column = sArrayRange(Column, sReshape(0, N, 1))
              'can be sure that Column is 1-based now
27            For i = 1 To N
28                If VarType(Column(i, 1)) = vbString Then
29                    Column(i, 2) = UCase$(Column(i, 1))
30                Else
31                    Column(i, 2) = Column(i, 1)
32                End If
33            Next i
34        Else
35            ColumnNumberForFirstSort = 1
36        End If

37        If Not SortAsWell Then
38            TempArray = ThrowIfError(sSortedArray(sArrayRange(Column, sIntegers(N)), ColumnNumberForFirstSort, , , , , , CaseSensitive))
39        Else
40            TempArray = ThrowIfError(sSortedArray(Column, ColumnNumberForFirstSort, , , , , , CaseSensitive))
41        End If

42        ReDim ChooseVector(1 To N, 1 To 1)
43        ChooseVector(1, 1) = True
44        For i = 2 To N
45            ChooseVector(i, 1) = Not sEquals(TempArray(i, 1), TempArray(i - 1, 1), CaseSensitive)
46        Next i
47        TempArray = sMChoose(TempArray, ChooseVector)

48        If Not SortAsWell Then
              'Have to put back into original order
49            TempArray = sSortedArray(TempArray, sNCols(TempArray), , , True)
50        End If

51        If sNCols(TempArray) > 1 Then
52            TempArray = sSubArray(TempArray, 1, 1, , 1)
53        End If

54        sRemoveDuplicates = TempArray

55        Exit Function
ErrHandler:
56        sRemoveDuplicates = "#sRemoveDuplicates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sRemoveDuplicatesColl
' Author     : Philip Swannell
' Date       : 19-May-2019
' Purpose    : More experimental code. Use a collection to remove duplicates. Key must be a string, hence need for method encode (which is tricksy)
'              hard to implement case-sensitivity since keys in collections are case-insensitive
' Parameters :
'  Column       :
'  SortAsWell   :
'  CaseSensitive:
' -----------------------------------------------------------------------------------------------------------------------
Function sRemoveDuplicatesColl(ByVal Column As Variant, _
        Optional SortAsWell As Boolean, _
        Optional CaseSensitive As Boolean = False)
          
          Dim c As Collection
          Dim EN As Long
          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result As Variant
          Dim STK As clsStacker
        
1         On Error GoTo ErrHandler
2         Force2DArrayR Column, NR, NC
3         If NC <> 1 Then Throw "Column must be an array of data with only one column"
        
4         Set STK = CreateStacker
5         Set c = New Collection

6         For i = 1 To NR
7             On Error Resume Next
8             c.Add 0, Encode(Column(i, 1), CaseSensitive)
9             EN = Err.Number
10            On Error GoTo ErrHandler
11            If EN = 0 Then
12                STK.Stack0D Column(i, 1)
13            End If
14        Next i
15        On Error GoTo ErrHandler
          
16        Result = STK.Report
17        If SortAsWell Then
18            Result = sSortedArray(Result, , , , , , , CaseSensitive)
19        End If
20        sRemoveDuplicatesColl = Result

21        Exit Function
ErrHandler:
22        sRemoveDuplicatesColl = "#sRemoveDuplicatesColl (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveDuplicatesXL
' Author    : Philip Swannell
' Date      : 13-Oct-2016
' Purpose   : Faster than sRemoveDuplicates for large arrays, but can't be called from
'             worksheet function.
' -----------------------------------------------------------------------------------------------------------------------
Private Function RemoveDuplicatesXL(Column, Optional SortAsWell As Boolean)
          Dim NR As Long
          Dim R As Range
          Dim SUH As clsScreenUpdateHandler
          Dim x As Long

1         On Error GoTo ErrHandler

2         NR = sNRows(Column)
3         If NR = 1 Then
4             RemoveDuplicatesXL = Column
5             Exit Function
6         End If

7         Set SUH = CreateScreenUpdateHandler()
8         shEmptySheet.UsedRange.EntireRow.Delete
9         Set R = shEmptySheet.Cells(1, 1).Resize(NR)
10        R.Value2 = sArrayExcelString(Column) 'Ensure that strings remain strings

          'The code below guards against a flaw in the .RemoveDuplicates method - that it does not distinguish between
          ' text "TRUE" and logical TRUE. Not sure if there are other examples, it does distinguish "1" from 1.
          ' R.Offset(, 1).FormulaR1C1 = "=IFS(ISNUMBER(RC[-1]),1,ISTEXT(RC[-1]),2,ISLOGICAL(RC[-1]),3,ISERROR(RC[-1]),4,ISBLANK(RC[-1]),5)"
11        R.Offset(, 1).FormulaR1C1 = "=IF(ISTEXT(RC[-1]),1,0)" 'should be sufficient to distinguish text and not-text
12        R.Resize(, 2).RemoveDuplicates Columns:=VBA.Array(1, 2), header:=xlNo 'Using VBA.Array ensures that the lower bound is zero, irrespective of Option Base

13        Set R = sExpandDown(R.Cells(1, 1))
14        If SortAsWell Then
15            R.Parent.Sort.SortFields.Clear
16            R.Parent.Sort.SortFields.Add Key:=R, SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
17            With R.Parent.Sort
18                .SetRange R
19                .header = xlNo
20                .MatchCase = False
21                .Orientation = xlTopToBottom
22                .SortMethod = xlPinYin
23                .Apply
24            End With
25        End If

26        RemoveDuplicatesXL = R.Value
27        If R.Rows.Count = 1 Then
28            Force2DArray RemoveDuplicatesXL        'for compatibility with sRemoveDuplicates
29        End If

30        shEmptySheet.UsedRange.EntireRow.Delete
31        x = shEmptySheet.UsedRange.Rows.Count        'This line resets the UsedRange

32        Exit Function
ErrHandler:
33        Throw "#RemoveDuplicatesXL (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Author     : Philip Swannell
' Date       : 18-May-2019
' Purpose    : Convert a variant to string in such a way that
' If CaseSensitive
'     UCase(Encode(x)) = UCase(Encode(y)) if and only if x = y
' If NOT CaseSensitive
'    UCase(Encode(x)) = UCase(Encode(y)) if and only if UCase(x) = UCase(y)
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(Optional x As Variant, Optional CaseSensitive As Boolean) As String
          Dim asciicodes As String
          Dim i As Long
          Dim strx As String
          Dim y As Double
1         strx = CStr(x)
2         Select Case VarType(x)
              Case vbDouble, vbSingle
3                 y = CDbl(strx)
4                 If x = y Then
5                     Encode = "5|" & strx 'In this case don't add third part, the string representation is exact. Thus we will recognise that (say) 1 = 1#
6                 Else
                      'will distinguish two unequal-but-very-nearly-equal doubles whose representation as strings (via CStr()is the same.
7                     Encode = "5|" & strx & "|" & CStr(x - y)
8                 End If
9             Case vbLong, vbInteger, vbSingle
10                Encode = "5|" & strx ' this way Long 5 and Double 5 will be taken as equal since their encoding will be the same
11            Case vbString
12                If CaseSensitive Then 'This is the tricky case - we need to make a case insensitive comparison of Encode(x) and Encode(y) mimic a case-sensitive comparison of x and y
13                    asciicodes = String(2 + 3 * Len(strx), " ")
14                    Mid$(asciicodes, 1, 2) = "3|"
15                    For i = 1 To Len(x)
16                        Mid$(asciicodes, 3 * i, 3) = Format$(Asc(Mid(strx, i, 1)), "000")
17                    Next
18                    Encode = asciicodes
19                Else
20                    Encode = "3|" & strx
21                End If
22            Case Else
23                Encode = CStr(VarType(x)) & "|" & strx
24        End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sRemoveDuplicatesDict
' Author     : Philip Swannell
' Date       : 19-May-2019
' Purpose    : Experimental code - use a dictonary to remove duplicates. Need to test performance - which seems not to be order n
'              but in my tests was faster than sort algorithm when Column was return from sIntegers for n up to 700,000.
'              The data was hard to extrapolate to get cross-over point for when sort might be faster.
' -----------------------------------------------------------------------------------------------------------------------
Function sRemoveDuplicatesDict(ByVal Column As Variant, _
        Optional SortAsWell As Boolean, _
        Optional CaseSensitive As Boolean = False)
          Dim DOther As Scripting.Dictionary
          Dim DStrings As Scripting.Dictionary
          Dim EN As Long
          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result As Variant
          Dim STK As clsStacker
        
1         On Error GoTo ErrHandler
2         Force2DArrayR Column, NR, NC
3         If NC <> 1 Then Throw "Column must be an array of data with only one column"
        
4         Set STK = CreateStacker
5         Set DStrings = New Scripting.Dictionary
6         If CaseSensitive Then
7             DStrings.CompareMode = BinaryCompare
8         Else
9             DStrings.CompareMode = TextCompare
10        End If
11        Set DOther = New Scripting.Dictionary

12        For i = 1 To NR
13            On Error Resume Next
14            If VarType(Column(i, 1)) = vbString Then
15                DStrings.Add Column(i, 1), 0
16                EN = Err.Number
17            Else
18                DOther.Add Column(i, 1), 0
19                EN = Err.Number
20            End If
21            On Error GoTo ErrHandler
22            If EN = 0 Then
23                STK.Stack0D Column(i, 1)
24            End If
25        Next i
26        On Error GoTo ErrHandler
          
27        Result = STK.Report
28        If SortAsWell Then
29            Result = sSortedArray(Result, , , , , , , CaseSensitive)
30        End If
31        sRemoveDuplicatesDict = Result

32        Exit Function
ErrHandler:
33        sRemoveDuplicatesDict = "#sRemoveDuplicatesDict (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sReshape
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Change shape of Input Array so that the output shape has NumRows and NumCols. Entries will
'             be read "like a book" line-by-line across the rows of Array starting from the
'             top and repeating as necessary to fill the new array.
' Arguments
' InputArray: An array of arbitrary values.
' NumRows   : The number of rows in the output array.
' NumCols   : The number of columns in the output array.
' -----------------------------------------------------------------------------------------------------------------------
Function sReshape(ByVal InputArray As Variant, Optional ByVal NumRows As Long, Optional ByVal NumCols As Long)
Attribute sReshape.VB_Description = "Change shape of Input Array so that the output shape has NumRows and NumCols. Entries will be read ""like a book"" line-by-line across the rows of Array starting from the top and repeating as necessary to fill the new array."
Attribute sReshape.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim a As Long
          Dim b As Long
          Dim i As Long
          Dim LB1 As Long
          Dim LB2 As Long
          Dim NumColsInput As Long
          Dim NumRowsInput As Long
          Dim NumElementsInput As Long
          Dim OutputArray() As Variant
          Dim x As Long
          Dim y As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR InputArray
3         LB1 = LBound(InputArray, 1)
4         LB2 = LBound(InputArray, 2)

5         NumRowsInput = UBound(InputArray, 1) - LB1 + 1
6         NumColsInput = UBound(InputArray, 2) - LB2 + 1
7         NumElementsInput = NumRowsInput * NumColsInput

8         If NumRows < 0 Then Throw "NumRows must be positive"
9         If NumCols < 0 Then Throw "NumCols must be positive"
10        If NumRows = 0 And NumCols = 0 Then Throw "At least one of NumRows and NumCols must be supplied"
11        If NumRows = 0 Then
12            If NumCols Mod NumElementsInput = 0 Then
13                NumRows = 1
14            ElseIf (NumElementsInput) Mod NumCols = 0 Then
15                NumRows = NumElementsInput / NumCols
16            Else
17                Throw "If NumRows is omitted then NumCols must be a factor or a multiple of the number of elements in InputArray"
18            End If
19        End If
20        If NumCols = 0 Then
21            If NumRows Mod NumElementsInput = 0 Then
22                NumCols = 1
23            ElseIf (NumElementsInput) Mod NumRows = 0 Then
24                NumCols = NumElementsInput / NumRows
25            Else
26                Throw "If NumCols is omitted then NumRows must be a factor or a multiple of the number of elements in InputArray"
27            End If
28        End If

29        ReDim OutputArray(1 To NumRows, 1 To NumCols)

30        a = 1: b = 1: x = 1: y = 1
31        For i = 1 To NumRows * NumCols
32            OutputArray(x, y) = InputArray(LB1 + a - 1, LB2 + b - 1)
33            If y = NumCols Then
34                y = 1
35                x = x + 1
36            Else
37                y = y + 1
38            End If
39            If b = NumColsInput Then
40                b = 1
41                If a = NumRowsInput Then
42                    a = 1
43                Else
44                    a = a + 1
45                End If
46            Else
47                b = b + 1
48            End If
49        Next i
50        sReshape = OutputArray

51        Exit Function
ErrHandler:
52        sReshape = "#sReshape (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowConcatenateStrings
' Author    : Philip Swannell
' Date      : 18-May-2015
' Purpose   : The function takes a row of strings and concatenates the entries together, with a
'             Delimiter string between each pair, to give a single output string. This is
'             the transpose version of sConcatenateStrings.
' Arguments
' TheStrings: An array of strings. If this array contains non-strings, then those elements will be cast
'             to strings before concatenation is done.
' Delimiter : The delimiter character. If omitted defaults to a comma. Can be specified as multiple
'             characters or the empty string.
' -----------------------------------------------------------------------------------------------------------------------
Function sRowConcatenateStrings(ByVal TheStrings As Variant, Optional Delimiter As String = ",") As Variant
Attribute sRowConcatenateStrings.VB_Description = "The function takes a row of strings and concatenates the entries together, with a Delimiter string between each pair, to give a single output string. This is the transpose version of sConcatenateStrings."
Attribute sRowConcatenateStrings.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim ResultArray() As String
          Dim TempArray() As String

1         On Error GoTo ErrHandler

2         Force2DArrayR TheStrings, NR, NC

3         ReDim ResultArray(1 To NR, 1 To 1)
4         ReDim TempArray(1 To NC)

5         For i = 1 To NR
6             For j = 1 To NC
7                 If VarType(TheStrings(i, j)) <> vbString Then
8                     TheStrings(i, j) = NonStringToString(TheStrings(i, j))
9                 End If
10                TempArray(j) = TheStrings(i, j)
11            Next j
12            ResultArray(i, 1) = VBA.Join(TempArray, Delimiter)
13        Next i

14        If NR = 1 Then
15            sRowConcatenateStrings = ResultArray(1, 1)
16        Else
17            sRowConcatenateStrings = ResultArray
18        End If

19        Exit Function
ErrHandler:
20        sRowConcatenateStrings = "#sRowConcatenateStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowMChoose
' Author    : Philip Swannell
' Date      : 28-Sep-2015
' Purpose   : The input ChooseVector must be a row array of logicals with the same width as TheArray.
'             The function returns those columns of TheArray for which the corresponding
'             flag in ChooseVector is TRUE. If no flags are TRUE, the function returns an
'             error.
' Arguments
' TheArray  : An array of arbitrary values.
' ChooseVector: A row array of TRUE and FALSE with the same width as TheArray.
' -----------------------------------------------------------------------------------------------------------------------
Function sRowMChoose(ByVal TheArray As Variant, ByVal ChooseVector As Variant)
Attribute sRowMChoose.VB_Description = "The input ChooseVector must be a row array of logicals with the same width as TheArray. The function returns those columns of TheArray for which the corresponding flag in ChooseVector is TRUE. If no flags are TRUE, the function returns an error."
Attribute sRowMChoose.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NCOut As Long
          Dim NR As Long
          Dim OutputArray As Variant
          Dim WriteCol As Long

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray, NR, NC
3         Force2DArrayR ChooseVector

          'Test Inputs
4         If UBound(ChooseVector, 2) <> NC Then
5             sRowMChoose = "#ChooseVector and TheArray must have the same number of columns!"
6             Exit Function
7         ElseIf UBound(ChooseVector, 1) <> 1 Then
8             sRowMChoose = "#ChooseVector must have one row only!"
9             Exit Function
10        End If

11        For i = 1 To UBound(ChooseVector, 2)
12            If VarType(ChooseVector(1, i)) = vbBoolean Then
13                If ChooseVector(1, i) Then
14                    NCOut = NCOut + 1
15                End If
16            Else
17                sRowMChoose = "#ChooseVector must contain only True or False!"
18                Exit Function
19            End If
20        Next i

21        If NCOut = 0 Then
22            sRowMChoose = "#Nothing to include!"
23            Exit Function
24        End If

25        ReDim OutputArray(1 To NR, 1 To NCOut)

26        For j = 1 To UBound(ChooseVector, 2)
27            If ChooseVector(1, j) Then
28                WriteCol = WriteCol + 1
29                For i = 1 To NR
30                    OutputArray(i, WriteCol) = TheArray(i, j)
31                Next i
32            End If
33        Next j

34        sRowMChoose = OutputArray

35        Exit Function
ErrHandler:
36        sRowMChoose = "#sRowMChoose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowReverse
' Author    : Philip Swannell
' Date      : 20-Jun-2013
' Purpose   : Flips an array left to right. The function returns an array of the same size as the input
'             array with the horizontal order reversed, so that the last column is now
'             first, and the first column is now last.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sRowReverse(TheArray As Variant)
Attribute sRowReverse.VB_Description = "Flips an array left to right. The function returns an array of the same size as the input array with the horizontal order reversed, so that the last column is now first, and the first column is now last."
Attribute sRowReverse.VB_ProcData.VB_Invoke_Func = " \n24"
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
4         Left = LBound(TheArray, 2)

5         ReDim ResultArray(1 To NR, 1 To NC)
6         For i = 1 To NR
7             For j = 1 To NC
8                 ResultArray(i, j) = TheArray(i - Top + 1, NC - j + Left)
9             Next j
10        Next i

11        sRowReverse = ResultArray

12        Exit Function
ErrHandler:
13        sRowReverse = "#sRowReverse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSubArray
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Extracts a rectangular subarray from TheArray. The subarray starts at the element indexed
'             by StartRow and StartColumn and the subarray has height Height and width
'             Width.
' Arguments
' TheArray  : The array from which a smaller "sub-array" is to be extracted.
' StartRow  : The index of the row of TheArray at which the subarray starts. Row indexing starts at 1.
' StartColumn: The index of the column of TheArray at which the subarray starts. Column indexing starts
'             at 1.
' Height    : The number of rows in the subarray.
' Width     : The number of columns in the subarray.
' -----------------------------------------------------------------------------------------------------------------------
Function sSubArray(TheArray, _
        Optional ByVal StartRow As Long = 1, _
        Optional ByVal StartColumn As Long = 1, _
        Optional ByVal Height As Long, _
        Optional ByVal Width As Long)
Attribute sSubArray.VB_Description = "Extracts a rectangular subarray from TheArray. The subarray starts at the element indexed by StartRow and StartColumn and the subarray has height Height and width Width.\n"
Attribute sSubArray.VB_ProcData.VB_Invoke_Func = " \n24"
          
          Dim i As Long
          Dim inFirstCol As Long
          Dim inFirstRow As Long
          Dim inHeight As Long
          Dim inWidth As Long
          Dim j As Long
          Dim ResultArray As Variant

1         On Error GoTo ErrHandler

2         Force2DArrayR TheArray, inHeight, inWidth

          'Morph omitted or negative inputs
3         If StartRow < 0 Then
4             StartRow = inHeight + StartRow + 1
5         End If
6         If StartColumn < 0 Then
7             StartColumn = inWidth + StartColumn + 1
8         End If
9         If Height = 0 Then Height = inHeight - StartRow + 1
10        If Width = 0 Then Width = inWidth - StartColumn + 1

          'Check inputs
11        If StartRow < 1 Or StartRow > inHeight Then
12            sSubArray = "#StartRow must be between 1 and " + CStr(inHeight) + _
                  " with negative values also allowed so as to count from the bottom of TheArray!"
13            Exit Function
14        ElseIf StartColumn < 1 Or StartColumn > inWidth Then
15            sSubArray = "#StartColumn must be between 1 and " + CStr(inWidth) + _
                  " with negative values also allowed so as to count from the right of TheArray!"
16            Exit Function
17        ElseIf Height < 1 Or StartRow + Height - 1 > inHeight Then
18            sSubArray = "#Invalid Height!"
19            Exit Function
20        ElseIf Width < 1 Or StartColumn + Width - 1 > inWidth Then
21            sSubArray = "#Invalid Width!"
22            Exit Function
23        End If

24        ReDim ResultArray(1 To Height, 1 To Width)

25        inFirstRow = LBound(TheArray, 1)
26        inFirstCol = LBound(TheArray, 2)

27        For i = 1 To Height
28            For j = 1 To Width
29                ResultArray(i, j) = TheArray(inFirstRow + StartRow - 2 + i, inFirstCol + StartColumn - 2 + j)
30            Next j
31        Next i

32        sSubArray = ResultArray

33        Exit Function
ErrHandler:
34        sSubArray = "#sSubArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSubRange
' Author    : Philip Swannell
' Date      : 13-Feb-2020
' Purpose   : Returns a sub-part of a range. Like Excel's OFFSET function, but different (often more
'             useful) behaviour for arguments 3 and 4 omitted.
' Arguments
' R         : A range of cells.
' FirstRow  : The index of the row of R at which the sub-range starts. Row indexing starts at 1.
'             Negative indexes can be used to count up from the bottom, i.e. -1 indicates
'             the last row. If omitted defaults to 1.
' FirstColumn: The index of the column of R at which the sub-range starts. Column indexing starts at 1.
'             Negative indexes can be used to left from the right, i.e. -1 indicates the
'             last column. If Omitted, defaults to 1.
' LastRow   : The index of the row of R at which the sub-range ends. Row indexing starts at 1. Negative
'             indexes can be used to count up from the bottom. If omitted defaults to -1
'             for the last row.
' LastColumn: The index of the column of R at which the sub-range ends. Column indexing starts at 1.
'             Negative indexes can be used to left from the right. If omitted defaults to
'             -1 for the last column.
' -----------------------------------------------------------------------------------------------------------------------
Function sSubRange(R As Range, Optional FirstRow = 1, Optional FirstColumn = 1, Optional LastRow = -1, Optional LastColumn = -1) As Range
Attribute sSubRange.VB_Description = "Returns a sub-part of a range. Like Excel's OFFSET function, but different (often more useful) behaviour for arguments 3 and 4 omitted."
Attribute sSubRange.VB_ProcData.VB_Invoke_Func = " \n31"
1         If FirstRow < 0 Then
2             FirstRow = R.Rows.Count + FirstRow + 1
3         End If
4         If FirstColumn < 0 Then
5             FirstColumn = R.Columns.Count + FirstColumn + 1
6         End If
7         If LastRow < 0 Then
8             LastRow = R.Rows.Count + LastRow + 1
9         End If
10        If LastColumn < 0 Then
11            LastColumn = R.Columns.Count + LastColumn + 1
12        End If
13        Set sSubRange = R.Offset(FirstRow - 1, FirstColumn - 1).Resize(LastRow - FirstRow + 1, LastColumn - FirstColumn + 1)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sTake
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns the first (or last) few rows of an array. If the absolute value of NumToTake is
'             greater than the height of TheArray then an error is returned
' Arguments
' TheArray  : An array of arbitrary values.
' NumToTake : Integer number of rows to take If positive, top rows are returned, and if negative, bottom
'             rows are returned.
' -----------------------------------------------------------------------------------------------------------------------
Function sTake(TheArray, NumToTake As Long)
Attribute sTake.VB_Description = "Returns the first (or last) few rows of an array. If the absolute value of NumToTake is greater than the height of TheArray then an error is returned"
Attribute sTake.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim NR As Long
1         NR = sNRows(TheArray)
2         If NumToTake = 0 Or NumToTake > NR Or NumToTake < -NR Then
3             sTake = "#Cannot take that many!"
4         ElseIf NumToTake > 0 Then
5             sTake = sSubArray(TheArray, , , NumToTake)
6         Else
7             sTake = sSubArray(TheArray, NR + NumToTake + 1, , -NumToTake)
8         End If
End Function

Sub TestRemoveDuplicatesXL()
1         On Error GoTo ErrHandler
          ' RemoveDuplicatesXL(sArrayStack(1, 1, 1, True, "TRUE", sIntegers(10), sIntegers(10)), False)
2         RemoveDuplicatesXL sArrayStack(1, 1, 1), False

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestRemoveDuplicatesXL (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : To1D
' Author     : Philip Swannell
' Date       : 24-Jan-2018
' Purpose    : Transform zero, one or higher dimensional array to one dimension. In the case of 2d, we read column-wise "top to bottom, left to right",
'              and for 3 or higher dimensions we read with the first dimension being on the inner-most loop. So an array A(i,j,k)
'              with i,j and k ranging from 1 to 3 would appear in the output in the following order:
'              {1,1,1}
'              {2,1,1}
'              {3,1,1}
'              {1,2,1}
'              {2,2,1}
'              {3,2,1}
'              {1,3,1}
'              {2,3,1}
'              {3,3,1}
'              {1,1,2}
'              {2,1,2}
'              {3,1,2}
'              {1,2,2}
'              {2,2,2}
'              etc.
' -----------------------------------------------------------------------------------------------------------------------
Function To1D(ByVal x As Variant)
          Dim i As Long
          Dim k As Long
          Dim ND As Long
          Dim Res() As Variant
1         On Error GoTo ErrHandler
2         If TypeName(x) = "Range" Then x = x.Value2
3         ND = NumDimensions(x)
4         Select Case ND
              Case 0
5                 ReDim Res(1 To 1)
6                 Res(1) = x
7                 To1D = Res
8             Case 1
9                 To1D = x
10            Case Else
                  Dim nOut As Long
                  Dim y As Variant
11                nOut = 1
12                For i = 1 To ND
13                    nOut = nOut * (UBound(x, i) - LBound(x, i) + 1)
14                Next i
15                k = 0
16                ReDim Res(1 To nOut)
17                For Each y In x
18                    k = k + 1
19                    Res(k) = y
20                Next
21                To1D = Res
22        End Select
23        Exit Function
ErrHandler:
24        Throw "#To1D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
