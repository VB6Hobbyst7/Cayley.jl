Attribute VB_Name = "modExpandRange"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpandDown
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended down until all cells just
'             underneath the returned range are empty, or until the range extends to the
'             last row of the worksheet.
' Arguments
' InputRange: A range of cells
' -----------------------------------------------------------------------------------------------------------------------
Function sExpandDown(InputRange As Range) As Range
Attribute sExpandDown.VB_Description = "Returns a reference to a range which is the InputRange extended down until all cells just underneath the returned range are empty, or until the range extends to the last row of the worksheet."
Attribute sExpandDown.VB_ProcData.VB_Invoke_Func = " \n31"
          Dim EarlyExit As Boolean
          Dim NonBlankCell As Range
          Dim NR As Long
          Dim Res

          Dim BottomRow As Range

1         On Error GoTo ErrHandler

2         Res = InputRange.Parent.UsedRange.Rows.Count        'Resets the UsedRange

3         With InputRange
4             If .Areas.Count > 1 Then Throw "InputRange must have only one area"
5             NR = .Worksheet.Rows.Count
6             If .row + .Rows.Count - 1 = NR Then
7                 Set sExpandDown = InputRange
8                 Exit Function
9             End If
10        End With

11        Set BottomRow = InputRange.Rows(InputRange.Rows.Count + 1)

12        Do Until IsRangeEmpty(BottomRow, NonBlankCell)
13            If BottomRow.row = NR Then
14                EarlyExit = True
15                Exit Do
16            End If
17            Set BottomRow = BottomRow.Offset(FindFirstBlank(NonBlankCell, xlDown).row - BottomRow.row)
18        Loop

19        Set sExpandDown = InputRange.Parent.Range(InputRange.Cells(1, 1), BottomRow.Cells(IIf(EarlyExit, 1, 0), BottomRow.Columns.Count))

20        Exit Function
ErrHandler:
          'Function returns a Range object, so we cannot return an error string
21        Throw "#sExpandDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpandUp
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended up until all cells just
'             above the returned range are empty, or until the range extends to the first
'             row of the worksheet.
' Arguments
' InputRange: A range of cells
' -----------------------------------------------------------------------------------------------------------------------
Function sExpandUp(InputRange As Range) As Range
Attribute sExpandUp.VB_Description = "Returns a reference to a range which is the InputRange extended up until all cells just above the returned range are empty, or until the range extends to the first row of the worksheet."
Attribute sExpandUp.VB_ProcData.VB_Invoke_Func = " \n31"
          Dim EarlyExit As Boolean
          Dim NonBlankCell As Range
          Dim Res

          Dim TopRow As Range

1         On Error GoTo ErrHandler

2         Res = InputRange.Parent.UsedRange.Rows.Count        'Resets the UsedRange
3         With InputRange
4             If .Areas.Count > 1 Then Throw "InputRange must have only one area"
5             If .row = 1 Then
6                 Set sExpandUp = InputRange
7                 Exit Function
8             End If
9         End With

10        Set TopRow = InputRange.Rows(0)

11        Do Until IsRangeEmpty(TopRow, NonBlankCell)
12            If TopRow.row = 1 Then
13                EarlyExit = True
14                Exit Do
15            End If
16            Set TopRow = TopRow.Offset(FindFirstBlank(NonBlankCell, xlUp).row - TopRow.row)
17        Loop

18        Set sExpandUp = InputRange.Parent.Range(TopRow.Cells(IIf(EarlyExit, 1, 2), 1), InputRange.Cells(InputRange.Rows.Count, InputRange.Columns.Count))

19        Exit Function
ErrHandler:
          'Function returns a Range object, so we cannot return an error string
20        Throw "#sExpandUp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpandRight
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended right until all cells just
'             to the right of the returned range are empty, or until the returned range
'             extends to the last column of the worksheet.
' Arguments
' InputRange: A range of cells
' -----------------------------------------------------------------------------------------------------------------------
Function sExpandRight(InputRange As Range) As Range
Attribute sExpandRight.VB_Description = "Returns a reference to a range which is the InputRange extended right until all cells just to the right of the returned range are empty, or until the returned range extends to the last column of the worksheet."
Attribute sExpandRight.VB_ProcData.VB_Invoke_Func = " \n31"
          Dim EarlyExit As Boolean
          Dim NC As Long
          Dim NonBlankCell As Range
          Dim Res
          Dim RightCol As Range

1         On Error GoTo ErrHandler
2         Res = InputRange.Parent.UsedRange.Rows.Count        'Resets the UsedRange
3         With InputRange
4             If .Areas.Count > 1 Then Throw "InputRange must have only one area"
5             NC = .Worksheet.Columns.Count
6             If .Column + .Columns.Count - 1 = NC Then
7                 Set sExpandRight = InputRange
8                 Exit Function
9             End If
10        End With

11        Set RightCol = InputRange.Columns(InputRange.Columns.Count + 1)

12        Do Until IsRangeEmpty(RightCol, NonBlankCell)
13            If RightCol.Column = NC Then
14                EarlyExit = True
15                Exit Do
16            End If
17            Set RightCol = RightCol.Offset(, FindFirstBlank(NonBlankCell, xlToRight).Column - RightCol.Column)
18        Loop

19        Set sExpandRight = InputRange.Parent.Range(InputRange.Cells(1, 1), RightCol.Cells(RightCol.Rows.Count, IIf(EarlyExit, 1, 0)))

20        Exit Function
ErrHandler:
          'Function returns a Range object, so we cannot return an error string
21        Throw "#sExpandRight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpandLeft
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended left until all cells just
'             to the left of the returned range are empty, or until the returned range
'             extends to the first column of the worksheet.
' Arguments
' InputRange: A range of cells
' -----------------------------------------------------------------------------------------------------------------------
Function sExpandLeft(InputRange As Range) As Range
Attribute sExpandLeft.VB_Description = "Returns a reference to a range which is the InputRange extended left until all cells just to the left of the returned range are empty, or until the returned range extends to the first column of the worksheet."
Attribute sExpandLeft.VB_ProcData.VB_Invoke_Func = " \n31"
          Dim EarlyExit As Boolean
          Dim LeftCol As Range
          Dim NonBlankCell As Range
          Dim Res

1         On Error GoTo ErrHandler

2         Res = InputRange.Parent.UsedRange.Rows.Count        'Resets the UsedRange

3         With InputRange
4             If .Areas.Count > 1 Then Throw "InputRange must have only one area"
5             If .Column = 1 Then
6                 Set sExpandLeft = InputRange
7                 Exit Function
8             End If
9         End With

10        Set LeftCol = InputRange.Columns(0)

11        Do Until IsRangeEmpty(LeftCol, NonBlankCell)
12            If LeftCol.Column = 1 Then
13                EarlyExit = True
14                Exit Do
15            End If
16            Set LeftCol = LeftCol.Offset(, FindFirstBlank(NonBlankCell, xlToLeft).Column - LeftCol.Column)
17        Loop

18        Set sExpandLeft = InputRange.Parent.Range(LeftCol.Cells(1, IIf(EarlyExit, 1, 2)), InputRange.Cells(InputRange.Rows.Count, InputRange.Columns.Count))

19        Exit Function
ErrHandler:
          'Function returns a Range object, so we cannot return an error string
20        Throw "#sExpandLeft (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpandRightDown
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended right and down until the
'             returned range is surrounded underneath and to the right by empty cells, or
'             until it extends to the last row or last column of the worksheet.
' Arguments
' InputRange: A range of cells
' -----------------------------------------------------------------------------------------------------------------------
Function sExpandRightDown(InputRange As Range) As Range
Attribute sExpandRightDown.VB_Description = "Returns a reference to a range which is the InputRange extended right and down until the returned range is surrounded underneath and to the right by empty cells, or until it extends to the last row or last column of the worksheet."
Attribute sExpandRightDown.VB_ProcData.VB_Invoke_Func = " \n31"
1         On Error GoTo ErrHandler
2         Set sExpandRightDown = sExpand(InputRange, False, True, False, True)
3         Exit Function
ErrHandler:
          'Function returns a Range object, so we cannot return an error string
4         Throw "#sExpandRightDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExpand
' Author    : Philip Swannell
' Date      : 20-Nov-2015
' Purpose   : Returns a reference to a range which is the InputRange extended in each of the requested
'             directions until cells bordering the returned range are empty.
' Arguments
' InputRange: A range of cells
' GoLeft    : True if the range is to be expanded leftwards. In this case all cells one to the left of
'             the returned range will be empty, or the returned range will extend to column
'             A.
' GoRight   : True if the range is to be expanded rightwards. In this case all cells one to the right of
'             the returned range will be empty, or the returned range will extend to the
'             last column of the worksheet.
' GoUp      : True if the range is to be expanded upwards. In this case all cells one above the returned
'             range will be empty, or the returned range will extend to row 1.
' GoDown    : True if the range is to be expanded downwards. In this case all cells one below the
'             returned range will be empty, or the returned range will extend to the last
'             row of the worksheet.
' -----------------------------------------------------------------------------------------------------------------------
Function sExpand(InputRange As Range, Optional GoLeft As Boolean = True, Optional goRight As Boolean = True, Optional GoUp As Boolean = True, Optional GoDown As Boolean = True) As Range
Attribute sExpand.VB_Description = "Returns a reference to a range which is the InputRange extended in each of the requested directions until cells bordering the returned range are empty. "
Attribute sExpand.VB_ProcData.VB_Invoke_Func = " \n31"
          Dim address As String
          Dim GoDownLeft As Boolean
          Dim GoDownRight As Boolean
          Dim GoUpLeft As Boolean
          Dim GoUpRight As Boolean
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim L As Long
          Dim NC As Long
          Dim NR As Long
          Dim OutPutRange As Range
          Dim Res

1         On Error GoTo ErrHandler
2         If Not (GoLeft Or goRight Or GoUp Or GoDown) Then
3             Set sExpand = InputRange
4             Exit Function
5         End If

6         Res = InputRange.Parent.UsedRange.Rows.Count        'Resets the UsedRange

7         GoDownRight = GoDown And goRight
8         GoDownLeft = GoDown And GoLeft
9         GoUpRight = GoUp And goRight
10        GoUpLeft = GoUp And GoLeft

11        Set OutPutRange = InputRange
12        Do
13            address = OutPutRange.address

14            If GoLeft Then Set OutPutRange = sExpandLeft(OutPutRange)
15            If goRight Then Set OutPutRange = sExpandRight(OutPutRange)
16            If GoUp Then Set OutPutRange = sExpandUp(OutPutRange)
17            If GoDown Then Set OutPutRange = sExpandDown(OutPutRange)

18            If Not (GoDownRight Or GoDownLeft Or GoUpRight Or GoUpLeft) Then
19                Set sExpand = OutPutRange
20                Exit Function
21            End If

22            With OutPutRange
23                If NR = 0 Then
24                    NR = InputRange.Parent.Rows.Count
25                    NC = InputRange.Parent.Columns.Count
26                End If
27                If GoDownRight Then
28                    i = 1        'i walks down right
TryAgainI:
29                    If .row + .Rows.Count - 1 + i <= NR Then
30                        If .Column + .Columns.Count - 1 + i <= NC Then
31                            If Not IsEmpty(.Cells(.Rows.Count + i, .Columns.Count + i)) Then
32                                i = i + 1
33                                GoTo TryAgainI
34                            End If
35                        End If
36                    End If
37                    i = i - 1
38                End If

39                If GoDownLeft Then
40                    j = 1        ' j walks down left
TryAgainJ:
41                    If .row + .Rows.Count - 1 + j <= NR Then
42                        If .Column - j >= 1 Then
43                            If Not IsEmpty(.Cells(.Rows.Count + j, 1 - j)) Then
44                                j = j + 1
45                                GoTo TryAgainJ
46                            End If
47                        End If
48                    End If
49                    j = j - 1
50                End If

51                If GoUpRight Then
52                    k = 1        ' k walks up right
TryAgainK:
53                    If .row - k >= 1 Then
54                        If .Column + .Columns.Count - 1 + k <= NC Then
55                            If Not IsEmpty(.Cells(1 - k, .Columns.Count + k)) Then
56                                k = k + 1
57                                GoTo TryAgainK
58                            End If
59                        End If
60                    End If
61                    k = k - 1
62                End If

63                If GoUpLeft Then
64                    L = 1        ' L walks up left
TryAgainL:
65                    If .row - L >= 1 Then
66                        If .Column - L >= 1 Then
67                            If Not IsEmpty(.Cells(1 - L, 1 - L)) Then
68                                L = L + 1
69                                GoTo TryAgainL
70                            End If
71                        End If
72                    End If
73                    L = L - 1
74                End If
75                If i > 0 Or j > 0 Or k > 0 Or L > 0 Then
76                    Set OutPutRange = .Offset(-SafeMax(L, k), -SafeMax(L, j)).Resize(.Rows.Count + SafeMax(L, k) + SafeMax(j, i), .Columns.Count + SafeMax(L, j) + SafeMax(k, i))
77                End If
78            End With
79        Loop While OutPutRange.address <> address

80        Set sExpand = OutPutRange

81        Exit Function
ErrHandler:
82        Throw "#sExpand (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FindFirstBlank
' Author    : Philip Swannell
' Date      : 18-Oct-2016
' Purpose   : c should be single cell range.
'             If c is empty returns c
'             Otherwise returns the first empty cell found when moving in the direction
'             specified (or the cell at the edge of the sheet if no blank cell is to be
'             found when moving in that direction).
'             NB as previous version (from Nov 2015) of this function made use of
'             .End(Direction) as a very fast way to seach for the last non-blank cell or
'             first blank cell. Unfortunately .End interacts badly with hidden rows and
'             columns - it behaves as if those rows/column have been deleted. Hence the
'             re-write of the code.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FindFirstBlank(c As Range, Direction As XlDirection) As Range
          Dim br As Long
          Dim Found As Boolean
          Dim i As Long
          Dim LC As Long
          Dim NC As Long
          Dim NR As Long
          Dim RangeToSearch As Range
          Dim RC As Long
          Dim TR As Long
          Dim UR As Range
          Dim ValuesToSearch
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         If IsEmpty(c.Value2) Then
3             Set FindFirstBlank = c
4             Exit Function
5         End If

6         Set ws = c.Parent
7         NR = ws.Rows.Count
8         NC = ws.Columns.Count
9         Set UR = ws.UsedRange
10        TR = UR.row
11        br = UR.row + UR.Rows.Count - 1
12        LC = UR.Column
13        RC = UR.Column + UR.Columns.Count - 1

14        If Application.Intersect(UR, c) Is Nothing Then Throw "Assertion Failed. Non blank cell is not in UsedRange"

15        Select Case Direction
              Case xlDown
16                Set RangeToSearch = ws.Range(c, ws.Cells(br, c.Column))
17                ValuesToSearch = RangeToSearch.Value2
18                If c.row = br Then Force2DArray ValuesToSearch
19                For i = 1 To br - c.row + 1
20                    If IsEmpty(ValuesToSearch(i, 1)) Then
21                        Found = True
22                        Exit For
23                    End If
24                Next i
25                If Found Then
26                    Set FindFirstBlank = c.Offset(i - 1)
27                ElseIf br < NR Then
28                    Set FindFirstBlank = ws.Cells(br + 1, c.Column)
29                Else
30                    Set FindFirstBlank = ws.Cells(NR, c.Column)
31                End If
32            Case xlUp
33                Set RangeToSearch = ws.Range(ws.Cells(TR, c.Column), c)
34                ValuesToSearch = RangeToSearch.Value2
35                If c.row = TR Then Force2DArray ValuesToSearch
36                For i = c.row - TR + 1 To 1 Step -1
37                    If IsEmpty(ValuesToSearch(i, 1)) Then
38                        Found = True
39                        Exit For
40                    End If
41                Next i
42                If Found Then
43                    Set FindFirstBlank = c.Offset(i - (c.row - TR + 1))
44                ElseIf TR > 1 Then
45                    Set FindFirstBlank = ws.Cells(TR - 1, c.Column)
46                Else
47                    Set FindFirstBlank = ws.Cells(1, c.Column)
48                End If
49            Case xlToRight
50                Set RangeToSearch = ws.Range(c, ws.Cells(c.row, RC))
51                ValuesToSearch = RangeToSearch.Value2
52                If c.Column = RC Then Force2DArray ValuesToSearch
53                For i = 1 To RC - c.Column + 1
54                    If IsEmpty(ValuesToSearch(1, i)) Then
55                        Found = True
56                        Exit For
57                    End If
58                Next i
59                If Found Then
60                    Set FindFirstBlank = c.Offset(, i - 1)
61                ElseIf RC < NC Then
62                    Set FindFirstBlank = ws.Cells(c.row, RC + 1)
63                Else
64                    Set FindFirstBlank = ws.Cells(c.row, NC)
65                End If
66            Case xlToLeft
67                Set RangeToSearch = ws.Range(ws.Cells(c.row, LC), c)
68                ValuesToSearch = RangeToSearch.Value2
69                If c.Column = LC Then Force2DArray ValuesToSearch
70                For i = c.Column - LC + 1 To 1 Step -1
71                    If IsEmpty(ValuesToSearch(1, i)) Then
72                        Found = True
73                        Exit For
74                    End If
75                Next i
76                If Found Then
77                    Set FindFirstBlank = c.Offset(, i - (c.Column - LC + 1))
78                ElseIf LC > 1 Then
79                    Set FindFirstBlank = ws.Cells(c.row, LC - 1)
80                Else
81                    Set FindFirstBlank = ws.Cells(c.row, 1)
82                End If
83        End Select
84        Exit Function
ErrHandler:
85        Throw "#FindFirstBlank (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsRangeEmpty
' Author    : Philip Swannell
' Date      : 07-Jan-2016
' Purpose   : Function tests if all cells of a range are empty. If not then function also
'             sets ByRef argument NonBlankFound. No guarantee that this is the first non empty cell (though it often will be)
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsRangeEmpty(RangeToTest As Range, Optional ByRef NonBlankFound As Range) As Boolean
          Dim c As Range
          Dim ConstantsRange As Range
          Dim D As Range
          Dim FormulasRange As Range
1         On Error GoTo ErrHandler

2         With RangeToTest
3             If .Cells.CountLarge < 70 Then        'For very small ranges it's fastest to loop around the cells
4                 For Each c In .Cells
5                     If Not IsEmpty(c.Value) Then
6                         Set NonBlankFound = c
7                         IsRangeEmpty = False
8                         Exit Function
9                     End If
10                Next c
11                IsRangeEmpty = True
12            ElseIf .Rows.Count = 1 Then        'sExpandRange only tests ranges with a single row or single column, _
                                                  fastest approach is to test the first cell and the EndRight (or EndDown) of the first cell
13                Set c = .Cells(1, 1)
14                If Not IsEmpty(c.Value) Then
15                    Set NonBlankFound = c
16                    IsRangeEmpty = False
17                    Exit Function
18                Else
19                    Set D = c.End(xlToRight)
20                    If D.Column > .Column + .Columns.Count - 1 Then
21                        IsRangeEmpty = True
22                        Exit Function
23                    ElseIf D.Column < .Column + .Columns.Count - 1 Then
24                        Set NonBlankFound = D
25                        IsRangeEmpty = False
26                        Exit Function
27                    Else        'it's possible that the end right took us to the far right of the worksheet
28                        If Not IsEmpty(D.Value) Then
29                            Set NonBlankFound = D
30                            IsRangeEmpty = False
31                            Exit Function
32                        Else
33                            IsRangeEmpty = True
34                            Exit Function
35                        End If
36                    End If
37                End If
38            ElseIf .Columns.Count = 1 Then
39                Set c = .Cells(1, 1)
40                If Not IsEmpty(c.Value) Then
41                    Set NonBlankFound = c
42                    IsRangeEmpty = False
43                    Exit Function
44                Else
45                    Set D = c.End(xlDown)
46                    If D.row > .row + .Rows.Count - 1 Then
47                        IsRangeEmpty = True
48                        Exit Function
49                    ElseIf D.row < .row + .Rows.Count - 1 Then
50                        Set NonBlankFound = D
51                        IsRangeEmpty = False
52                        Exit Function
53                    Else        'it's possible that the end right took us to the bottom row of the worksheet
54                        If Not IsEmpty(D.Value) Then
55                            Set NonBlankFound = D
56                            IsRangeEmpty = False
57                            Exit Function
58                        Else
59                            IsRangeEmpty = True
60                            Exit Function
61                        End If
62                    End If
63                End If
64            Else
                  'For larger ranges, it's faster (in the case of all cells blank) to use the .SpecialCells method
                  '.SpecialCells has "surprising" behaviour when passed a single cell (operates on entire worksheet!) but we know that's not the case
                  'Documentation for .SpecialCells suggests that you can add the two constants to use
                  'SpecialCells(xlCellTypeConstants + xlCellTypeFormulas) but that fails if there are no constants but some formulas
                  'so we use each constant in turn...
65                On Error Resume Next
66                Set ConstantsRange = .SpecialCells(xlCellTypeConstants)
67                On Error GoTo ErrHandler
68                If Not ConstantsRange Is Nothing Then
69                    IsRangeEmpty = False
70                    Set NonBlankFound = ConstantsRange.Areas(1).Cells(1, 1)
71                    Exit Function
72                End If

73                On Error Resume Next
74                Set FormulasRange = .SpecialCells(xlCellTypeFormulas)
75                On Error GoTo ErrHandler
76                If Not FormulasRange Is Nothing Then
77                    IsRangeEmpty = False
78                    Set NonBlankFound = FormulasRange.Areas(1).Cells(1, 1)
79                    Exit Function
80                Else
81                    IsRangeEmpty = True
82                End If
83            End If
84        End With
85        Exit Function
ErrHandler:
86        Throw "#IsRangeEmpty (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRangeContainsRange
' Author    : Philip Swannell
' Date      : 06-Jan-2016
' Purpose   : Returns TRUE if SmallRange is entirely inside BigRange, FALSE otherwise
' -----------------------------------------------------------------------------------------------------------------------
Function sRangeContainsRange(BigRange As Range, SmallRange As Range) As Boolean
1         On Error GoTo ErrHandler
2         If BigRange.Parent Is SmallRange.Parent Then
3             sRangeContainsRange = Application.Union(BigRange, SmallRange).Cells.CountLarge = BigRange.Cells.CountLarge
4         Else
5             sRangeContainsRange = False
6         End If
7         Exit Function
ErrHandler:
8         Throw "#sRangeContainsRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
