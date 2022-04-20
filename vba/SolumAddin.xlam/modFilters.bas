Attribute VB_Name = "modFilters"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FilterRangeByHidingRows
' Author    : Philip Swannell
' Date      : 26-Apr-2016
' Purpose   : An alternative to Excel's built-in data filtering.
'             Message is a ByRef argument and is set to e.g. "333 of 1,000 rows shown" caller
'             can post this message to the screen in some helpful way. RowDescriptor is used
'             in the construction of Message.
'             InputMessages can be passed as a column array, and sets the Data Validation. Alternatively pass as True for
'                 auto-generated input messages, or False for no messages.
'             InputMessage for the cells of the FilterRange
'             RegKeys should be a column array of strings, one associated with each column of the data
'             and are used to persist "most recently used" filters in the registry
' -----------------------------------------------------------------------------------------------------------------------
Sub FilterRangeByHidingRows(FilterRange As Range, DataRange As Range, Optional RowDescriptor = "row", _
          Optional ByRef Message As String, Optional ByVal InputMessages As Variant = True, Optional RegKeys As Variant)

          Dim anyExcl As Boolean
          Dim anyHidden As Boolean
          Dim anyIncl As Boolean
          Dim ChooseVector As Variant
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NumShown As Long
          Dim InputMessageType As Long ', 0 = none,1 = passed in , 2 = auto
          Dim RangeToShow As Range
          Dim rx As VBScript_RegExp_55.RegExp
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim ThisColValues As Variant
          Dim WriteToRegistry As Boolean
          Dim InputMessage As String
          Dim UseTextProperty As Boolean
          Dim UseCStr As Boolean
          Dim UseThisFormat As String

1         On Error GoTo ErrHandler

2         If FilterRange.Rows.Count <> 1 Then Throw "FilterRange must have one row only"
3         If FilterRange.Columns.Count <> DataRange.Columns.Count Then Throw "FilterRange and DataRange must have the same number of columns"
4         If Not IsMissing(InputMessages) Then
5             If VarType(InputMessages) = vbBoolean Then
6                 If InputMessages Then
7                     InputMessageType = 2
8                 Else
9                     InputMessageType = 0
10                End If

11            ElseIf sNRows(InputMessages) <> FilterRange.Columns.Count Then
12                Throw "InputMessages must be a column array of text  with the same number of elements as DataRange has columns, or else True to auto-generate input messages"
13            Else
14                InputMessageType = 1
15            End If
16        End If
17        If Not IsMissing(RegKeys) Then
18            If sNRows(RegKeys) <> FilterRange.Columns.Count Then
19                Throw "RegKeys must be a column array of text  with the same number of elements as DataRange has columns"
20            End If
21            WriteToRegistry = True
22        End If

23        Set SUH = CreateScreenUpdateHandler()
24        Set SPH = CreateSheetProtectionHandler(DataRange.Parent)

25        With FilterRange
26            .Locked = False
27            .NumberFormat = "@"
28            .HorizontalAlignment = xlHAlignCenter
29            .Font.Color = RGB(0, 0, 255)
30        End With

31        NR = DataRange.Rows.Count
32        NC = DataRange.Columns.Count
33        ChooseVector = sReshape(True, NR, 1)

34        Set rx = New RegExp
35        For j = 1 To NC

36            If InputMessageType = 1 Then
37                SetCellValidation FilterRange.Cells(1, j), True, CStr(InputMessages(j, 1)), ""
38            ElseIf InputMessageType = 2 Then
39                SetCellValidation FilterRange.Cells(1, j), True, _
                      vbLf & "Filter " & RowDescriptor & "s by entering text. Double-click for more options.", _
                      DataRange.Cells(0, j).Value
40            Else
41                SetCellValidation FilterRange.Cells(1, j), False, "", ""
42                InputMessage = vbNullString
43            End If
44        Next j

45        For j = 1 To NC
46            If VarType(FilterRange.Cells(1, j).Value) = vbString Then
47                If RegExSyntaxValid(FilterRange.Cells(1, j).Value) Then
48                    anyIncl = False: anyExcl = False

49                    With rx
50                        .IgnoreCase = True
51                        .Pattern = FilterRange.Cells(1, j).Value
52                        .Global = False        'Find first match only
53                    End With
54                    ThisColValues = DataRange.Columns(j).Value        'use Value not Value2
                      'Arrgh the .Text property is slow to call, try to avoid.
55                    TestToUseTextProperty DataRange.Columns(j), UseTextProperty, UseCStr, UseThisFormat

56                    If NR = 1 Then Force2DArray ThisColValues
57                    For i = 1 To NR
58                        If ChooseVector(i, 1) Then
59                            Select Case VarType(ThisColValues(i, 1))
                                  Case vbString
60                                    ChooseVector(i, 1) = rx.Test(ThisColValues(i, 1))
61                                Case vbEmpty
62                                    ChooseVector(i, 1) = rx.Test(vbNullString)
63                                Case vbDouble, vbInteger, vbSingle, vbLong, vbDate, vbCurrency, 20 '(20 = vbLongLong)
64                                    If UseTextProperty Then
65                                        ChooseVector(i, 1) = rx.Test(DataRange.Cells(i, j).text)
66                                    ElseIf UseCStr Then
67                                        ChooseVector(i, 1) = rx.Test(CStr(ThisColValues(i, 1)))
68                                    Else
69                                       ChooseVector(i, 1) = rx.Test(Format$(ThisColValues(i, 1), UseThisFormat))
70                                    End If
71                                Case vbError
72                                    ChooseVector(i, 1) = rx.Test(NonStringToString(ThisColValues(i, 1)))
73                                Case vbBoolean
74                                    ChooseVector(i, 1) = rx.Test(UCase$(CStr(ThisColValues(i, 1))))
75                                Case Else
76                                    ChooseVector(i, 1) = rx.Test(CStr(ThisColValues(i, 1)))        'this line should never be hit
77                            End Select
78                            If Not anyHidden Then        'anyHidden indicates whether any row is to be hidden across all columns of filters
79                                anyHidden = Not (ChooseVector(i, 1))
80                            End If
81                            If WriteToRegistry Then If ChooseVector(i, 1) Then anyIncl = True Else anyExcl = True
82                        End If
83                    Next i
                      ' only add a filter to the registry if it's useful i.e. there are some matches and some non-matches
84                    If WriteToRegistry Then If (anyIncl And anyExcl) Then AddFilterToMRU CStr(RegKeys(j, 1)), FilterRange.Cells(1, j).Value
85                Else
86                    InputMessage = "That's not a valid regular expression. Try double-clicking the cell, " & _
                          "then select 'Advanced Filtering` for a dialog to help construct simple and valid regular " & _
                          "expressions that will be used to filter the contents of this column."
87                    SetCellValidation FilterRange.Cells(1, j), True, InputMessage, "Oops"
88                End If
89            End If
90        Next j
            
91        If Not anyHidden Then
92            DataRange.EntireRow.Hidden = False
93            NumShown = NR
94        Else
95            sFilterRange DataRange, ChooseVector, RangeToShow
96            NumShown = sArrayCount(ChooseVector)
97            DataRange.EntireRow.Hidden = True
98            If Not RangeToShow Is Nothing Then
99                RangeToShow.EntireRow.Hidden = False
100           End If
101       End If

102       If NR = 1 Then
              Dim c As Range
103           NR = 0
104           For Each c In DataRange.Cells
105               If Not IsEmpty(c.Value) Then
106                   NR = 1
107                   Exit For
108               End If
109           Next c
110       End If

111       If NR = 0 Then
112           Message = "No " + RowDescriptor + "s to show."
113       ElseIf NumShown = NR Then
114           Message = "All " + Format$(NR, "###,##0") + " " + RowDescriptor + IIf(NR > 1, "s", vbNullString) + " shown."
115       Else
116           Message = Format$(NumShown, "###,##0") + " of " + Format$(NR, "###,##0") + " " + RowDescriptor + IIf(NR > 1, "s", vbNullString) + " shown."
117       End If

118       Exit Sub
ErrHandler:
119       Throw "#FilterRangeByHidingRows (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestToUseTextProperty
' Author     : Philip Swannell
' Date       : 18-Mar-2022
' Purpose    : Sub of both FilterRangeByHidingRows and ShowRegularExpressionDialog. Figures out whether we can avoid
'              using the rather slow .Text property of cells as the text against which regular expression matching is done,
'              for better speed we would rather use CStr or VBA.Format
' -----------------------------------------------------------------------------------------------------------------------
Sub TestToUseTextProperty(R As Range, ByRef UseTextProperty As Boolean, ByRef UseCStr As Boolean, UseThisFormat As String)
          Dim NumberFormat As Variant

1         On Error GoTo ErrHandler
2         NumberFormat = R.NumberFormat
3         If IsNull(NumberFormat) Then
4             UseTextProperty = True
5             UseCStr = False
6             UseThisFormat = ""
7             Exit Sub
8         End If
9         If NumberFormat = "General" Then
10            UseTextProperty = False
11            UseCStr = True
12            UseThisFormat = ""
13        Else
              'Test (rather a weak test) whether VBA's Format function matches Excel's NumberFormat.
14            With shEmptySheet.Cells(1, 1)
15                .Clear
16                .ColumnWidth = 8.43
17                .Value = 1000
18                .NumberFormat = NumberFormat
19                UseTextProperty = .text <> Format$(1000, NumberFormat)
20                .Clear
21            End With
22            If UseTextProperty Then
23                UseCStr = False
24                UseThisFormat = ""
25            Else
26                UseCStr = False
27                UseThisFormat = NumberFormat
28            End If
29        End If

30        Exit Sub
ErrHandler:
31        Throw "#TestToUseTextProperty (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetCellValidation
' Author     : Philip Swannell
' Date       : 10-Mar-2022
' Purpose    : Sub of FilterRangeByHidingRows. Set a cell's validation, written for speed i.e don't set properties if
'              they already have the value we require.
' -----------------------------------------------------------------------------------------------------------------------
 Private Sub SetCellValidation(cell As Range, DoValidation As Boolean, InputMessage As String, InputTitle As String)

          Dim HasValidation As Boolean
          Dim tmp As Variant
          Dim v As Validation
          Dim EN As Long

1         On Error Resume Next
2         tmp = cell.Validation.InputMessage
3         EN = Err.Number
4         On Error GoTo ErrHandler

5         HasValidation = EN = 0

6         If HasValidation Then
7             Set v = cell.Validation
8             If DoValidation Then
9                 If v.InputMessage <> InputMessage Then
10                    v.InputMessage = InputMessage
11                End If
12                If v.InputTitle <> InputTitle Then
13                    v.InputTitle = InputTitle
14                End If
15            Else
16                v.Delete
17            End If
18        Else
19            If DoValidation Then
20                cell.Validation.Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
21                Set v = cell.Validation
22                v.InputMessage = InputMessage
23                If InputTitle <> "" Then
24                    v.InputTitle = InputTitle
25                End If
26            End If
27        End If

28        Exit Sub
ErrHandler:
29        Throw "#SetCellValidation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFilterRange
' Author     : Philip Swannell
' Date       : 10-Mar-2022
' Purpose    : Like MChoose, but the return is a (multi-area) Range object. Can be called only from VBA. If no element of
'              ChooseVecor is True then the return is Nothing. This funciton cannot be called from a worksheet since
'              functions cannot return multiple-area ranges to a worksheet
' Parameters :
'  RangeIn    :
'  ChooseVector:
' -----------------------------------------------------------------------------------------------------------------------
Sub sFilterRange(RangeIn As Range, ChooseVector, ByRef RangeOut As Range)

          Dim NR As Long, NC As Long, i As Long
          Dim objF As clsStacker
          Dim objH As clsStacker
          Dim H As Long
          Dim FHArray
          Dim CountingTrues As Boolean
          Dim anyTrue As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayR ChooseVector, NR, NC

          'Validate inputs
3         If NR <> RangeIn.Rows.Count Then
4             Throw "ChooseVector must have the same number of rows as RangeIn"
5         End If

6         For i = 1 To NR
7             If VarType(ChooseVector(i, 1)) <> vbBoolean Then
8                 Throw "ChooseVector must contain Booleans only"
9             End If
10            If Not anyTrue Then If ChooseVector(i, 1) Then anyTrue = True
11        Next i

12        If Not anyTrue Then
13            Set RangeOut = Nothing
14            Exit Sub
15        End If
          'Construct FHArray, the key argument to sFRCore
16        Set objF = CreateStacker()
17        Set objH = CreateStacker()

18        H = 1

19        If ChooseVector(1, 1) Then
20            objF.Stack0D 1
21            CountingTrues = True
22        End If

23        For i = 2 To NR
24            If ChooseVector(i, 1) = ChooseVector(i - 1, 1) Then
25                H = H + 1
26            Else
27                If CountingTrues Then
28                    objH.Stack0D H
29                Else
30                    objF.Stack0D i
31                End If
32                CountingTrues = Not CountingTrues
33                H = 1
34            End If
35        Next i
36        If CountingTrues Then
37            objH.Stack0D H
38        End If

39        FHArray = sArrayRange(objF.Report, objH.Report)

40        Set RangeOut = sFRCore(RangeIn, FHArray, 1, sNRows(FHArray))

41        Exit Sub
ErrHandler:
42        Throw "#sFilterRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFRCore - core of sFilterRange
' Author     : Philip Swannell
' Date       : 09-Mar-2022
' Purpose    : Returns a range being the union of specified rows in the input range R. Uses a recursive algorithm for speed.
' Parameters :
'  R        : A Range
'  FHArray  : A two column array of "From Rows" and "How Many Rows" example if FHArray is {2,2;10,3} then that specifies
'             that the return should be rows 2,3,10,11,12
'  FromIndex: Only those rows of FHArray starting from row FromIndex are processed.
'  ToIndex  : Only those rows of FHArray ending at this index are processed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function sFRCore(R As Range, FHArray As Variant, FromIndex As Long, ToIndex As Long)

          Dim N As Long
          Dim mid1 As Long
          Dim mid2 As Long
          Dim mid3 As Long
          Dim mid4 As Long
          Dim mid5 As Long
          Dim mid6 As Long
          Dim mid7 As Long
          Dim mid8 As Long
          Dim mid9 As Long

1         On Error GoTo ErrHandler

2         N = ToIndex - FromIndex + 1

3         Select Case N
              Case 1
4                 Set sFRCore = R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2))
5             Case 2
6                 Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)))
7             Case 3
8                 Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)))
9             Case 4
10                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)))
11            Case 5
12                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)))
13            Case 6
14                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)), _
                      R.Rows(FHArray(FromIndex + 5, 1)).Resize(FHArray(FromIndex + 5, 2)))
15            Case 7
16                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)), _
                      R.Rows(FHArray(FromIndex + 5, 1)).Resize(FHArray(FromIndex + 5, 2)), _
                      R.Rows(FHArray(FromIndex + 6, 1)).Resize(FHArray(FromIndex + 6, 2)))
17            Case 8
18                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)), _
                      R.Rows(FHArray(FromIndex + 5, 1)).Resize(FHArray(FromIndex + 5, 2)), _
                      R.Rows(FHArray(FromIndex + 6, 1)).Resize(FHArray(FromIndex + 6, 2)), _
                      R.Rows(FHArray(FromIndex + 7, 1)).Resize(FHArray(FromIndex + 7, 2)))
19            Case 9
20                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)), _
                      R.Rows(FHArray(FromIndex + 5, 1)).Resize(FHArray(FromIndex + 5, 2)), _
                      R.Rows(FHArray(FromIndex + 6, 1)).Resize(FHArray(FromIndex + 6, 2)), _
                      R.Rows(FHArray(FromIndex + 7, 1)).Resize(FHArray(FromIndex + 7, 2)), _
                      R.Rows(FHArray(FromIndex + 8, 1)).Resize(FHArray(FromIndex + 8, 2)))
21            Case 10
22                Set sFRCore = Application.Union( _
                      R.Rows(FHArray(FromIndex, 1)).Resize(FHArray(FromIndex, 2)), _
                      R.Rows(FHArray(FromIndex + 1, 1)).Resize(FHArray(FromIndex + 1, 2)), _
                      R.Rows(FHArray(FromIndex + 2, 1)).Resize(FHArray(FromIndex + 2, 2)), _
                      R.Rows(FHArray(FromIndex + 3, 1)).Resize(FHArray(FromIndex + 3, 2)), _
                      R.Rows(FHArray(FromIndex + 4, 1)).Resize(FHArray(FromIndex + 4, 2)), _
                      R.Rows(FHArray(FromIndex + 5, 1)).Resize(FHArray(FromIndex + 5, 2)), _
                      R.Rows(FHArray(FromIndex + 6, 1)).Resize(FHArray(FromIndex + 6, 2)), _
                      R.Rows(FHArray(FromIndex + 7, 1)).Resize(FHArray(FromIndex + 7, 2)), _
                      R.Rows(FHArray(FromIndex + 8, 1)).Resize(FHArray(FromIndex + 8, 2)), _
                      R.Rows(FHArray(FromIndex + 9, 1)).Resize(FHArray(FromIndex + 9, 2)))
23            Case Is > 10
24                mid1 = FromIndex + N \ 10
25                mid2 = FromIndex + (2 * N) \ 10
26                mid3 = FromIndex + (3 * N) \ 10
27                mid4 = FromIndex + (4 * N) \ 10
28                mid5 = FromIndex + (5 * N) \ 10
29                mid6 = FromIndex + (6 * N) \ 10
30                mid7 = FromIndex + (7 * N) \ 10
31                mid8 = FromIndex + (8 * N) \ 10
32                mid9 = FromIndex + (9 * N) \ 10

33                Set sFRCore = Application.Union( _
                      sFRCore(R, FHArray, FromIndex, mid1), _
                      sFRCore(R, FHArray, mid1 + 1, mid2), _
                      sFRCore(R, FHArray, mid2 + 1, mid3), _
                      sFRCore(R, FHArray, mid3 + 1, mid4), _
                      sFRCore(R, FHArray, mid4 + 1, mid5), _
                      sFRCore(R, FHArray, mid5 + 1, mid6), _
                      sFRCore(R, FHArray, mid6 + 1, mid7), _
                      sFRCore(R, FHArray, mid7 + 1, mid8), _
                      sFRCore(R, FHArray, mid8 + 1, mid9), _
                      sFRCore(R, FHArray, mid9 + 1, ToIndex))
                      
34            Case Else
35                Throw "Error in recursive algorithm, variable N has the unexpected value of " & CStr(N)
36        End Select

37        Exit Function
ErrHandler:
38        Throw "#sFRCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

