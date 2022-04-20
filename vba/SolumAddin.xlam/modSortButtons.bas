Attribute VB_Name = "modSortButtons"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Methods to add and remove "sort buttons" from worksheets. Adding buttons,
'             removing buttons and sorting by clicking on a button are all undoable using
'             code triggered via Application.OnUndo.
' -----------------------------------------------------------------------------------------------------------------------

'Module-level variables needed to implement undo functionality
Private m_WsUndoAdd As Worksheet
Private m_wsUndoSort As Worksheet
Private m_WsUndoRemove As Worksheet
Private m_UndoAddList As Variant
Private m_UndoRemoveList As Variant
Private m_ButtonStates As Variant        ' Three element column array: ButtonRange.Address,NonFlatButton.Name,NonFlatButton.Direction
Private m_LastNumHeaderRows As Long
Private Const MAX_BUTTONS_FOR_ADD = 1000

'Module_level variables for achieving speed-up in method ResetSortButtons
Private m_LastButtonWithArrow As String

Option Explicit

Enum EnmsbDirection
    sbDirectionUp = 1
    sbDirectionFlat = 0
    sbDirectionDown = -1
End Enum

Private Function m_MsgBoxTitle() As String
1         m_MsgBoxTitle = "Sort Buttons (" & gAddinName & ")"
End Function

Public Sub RepeatAddSortButtons()
1         AddSortButtons , m_LastNumHeaderRows
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Method to add sort buttons to a range of cells (the top row of the current selection)
'             Rather than calling the method AddSortButton in a loop we call AddSortButton once and
'             then use Copy and Paste. For adding 1000 buttons this is 8 times faster (1.39 vs 11.44 seconds)
' -----------------------------------------------------------------------------------------------------------------------
Public Sub AddSortButtons(Optional RangeForButtons As Range, Optional NumHeaderRows As Long = 1)
1         On Error GoTo ErrHandler
          Dim c As Range
          Dim ExSH As clsExcelStateHandler
          Dim FirstButton As Button
          Dim RangeForAllButFirstButton As Range
          Dim SPH As clsSheetProtectionHandler
          Dim STK As clsStacker
          Dim SUH As clsScreenUpdateHandler
          Dim withUndo As Boolean

2         If IsMissing(RangeForButtons) Or RangeForButtons Is Nothing Then
3             If TypeName(Selection) <> "Range" Then Exit Sub
4             Set RangeForButtons = Selection.Areas(1).Rows(1)
5             withUndo = True
6         End If

7         Set SPH = CreateSheetProtectionHandler(RangeForButtons.Parent)
8         Set SUH = CreateScreenUpdateHandler()
9         Set ExSH = CreateExcelStateHandler(, , , , , True)        'avoid changing the viewport

10        If RangeForButtons.Cells.CountLarge > MAX_BUTTONS_FOR_ADD Then
11            Throw "You cannot add sort buttons to a range with more than " + CStr(MAX_BUTTONS_FOR_ADD) + " columns"
12        End If

13        If NumHeaderRows < 0 Then
              Dim BadRes As Boolean
              Dim ExtraText As String
              Dim Res As String
TryAgain:
14            Res = InputBoxPlus("How many header rows does the range to" + vbLf + "be sorted have?" + ExtraText, "Add Sort Buttons", "2")
15            If Res = "False" Then Exit Sub
16            BadRes = False
17            If Not IsNumeric(Res) Then
18                BadRes = True
19            ElseIf CLng(Res) < 0 Or CLng(Res) > 100 Then
20                BadRes = True
21            End If
22            If BadRes Then
23                ExtraText = " (Must be a number in the" + vbLf + "range 0 to 100)."
24                GoTo TryAgain
25            End If
26            NumHeaderRows = CLng(Res)
27        End If
          'We don't want there to be two sort buttons sitting on any cell so remove any that might be there already...
28        RemoveSortButtonsInRange RangeForButtons, False

29        If withUndo Then
30            Set STK = CreateStacker()
31        End If

32        AddSortButton FirstButton, RangeForButtons.Cells(1, 1), NumHeaderRows
33        If withUndo Then
34            STK.Stack0D FirstButton.Name
35        End If

          ' This approach of copying and pasting cells is faster but seems unreliable - sometimes yields error "Paste method of Worksheet class failed"
          '    If RangeForButtons.Cells.CountLarge > 1 Then
          '        Set RangeForAllButFirstButton = RangeForButtons.Offset(0, 1).Resize(, RangeForButtons.Columns.Count - 1)
          '        FirstButton.Copy
          '        FirstButtonWidth = FirstButton.Width
          '
          '        For Each c In RangeForAllButFirstButton.Cells
          '            Application.GoTo c
          '            RangeForButtons.Parent.Paste
          '            If withUndo Then
          '                STK.Stack0D Selection.Name
          '            End If
          '            If FirstButtonWidth <> c.Width Then
          '                Selection.Width = c.Width
          '            End If
          '        Next c
          '    End If

36        If RangeForButtons.Cells.CountLarge > 1 Then
37            Set RangeForAllButFirstButton = RangeForButtons.Offset(0, 1).Resize(, RangeForButtons.Columns.Count - 1)
38            For Each c In RangeForAllButFirstButton.Cells
                  Dim b As Button
39                AddSortButton b, c, NumHeaderRows
40                If withUndo Then
41                    STK.Stack0D b.Name
42                End If
43            Next c
44        End If
          
45        Application.GoTo RangeForButtons.Cells(1, 1)
46        m_LastNumHeaderRows = NumHeaderRows
          Dim RepeatText As String
47        RepeatText = "Repeat Add Sort Buttons ("
48        Select Case NumHeaderRows
              Case 0
49                RepeatText = RepeatText + "no header rows)"
50            Case 1
51                RepeatText = RepeatText + "one header rows)"
52            Case Else
53                RepeatText = RepeatText + CStr(NumHeaderRows) + " header rows)"
54        End Select
        
55        If withUndo Then
56            m_UndoAddList = STK.Report
57            Set m_WsUndoAdd = RangeForButtons.Parent
58            Application.GoTo RangeForButtons
59            Application.OnUndo "Undo Add Sort Buttons to " & AddressND(RangeForButtons), "UndoAddSortButtons"
60        End If
61        Application.OnRepeat RepeatText, "RepeatAddSortButtons"

62        Exit Sub
ErrHandler:
63        Throw "#AddSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UndoAddSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Implements undo for most recently added sort buttons!
' -----------------------------------------------------------------------------------------------------------------------
Public Sub UndoAddSortButtons()
          Dim i As Long
          Dim SheetName As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler

2         On Error Resume Next
3         SheetName = m_WsUndoAdd.Name
4         On Error GoTo ErrHandler

5         If SheetName = vbNullString Then Exit Sub
6         If Not IsArray(m_UndoAddList) Then Exit Sub

7         Set SPH = CreateSheetProtectionHandler(m_WsUndoAdd)
8         Application.ScreenUpdating = False
9         For i = 1 To sNRows(m_UndoAddList)
10            m_WsUndoAdd.Buttons(m_UndoAddList(i, 1)).Delete
11        Next i
12        m_UndoAddList = Empty

13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#UndoAddSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Method to remove sort buttons that intersect the current selection. Available from the Ribbon.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RemoveSortButtons()
          Dim N As Long
          Dim RangeofButtons As Range
          Dim SPH As clsSheetProtectionHandler
          Dim UndoText As String

1         On Error GoTo ErrHandler

2         If TypeName(Selection) = "Range" Then
3             Set SPH = CreateSheetProtectionHandler(ActiveSheet)
4             Application.ScreenUpdating = False
5             Set RangeofButtons = Selection
6             N = RemoveSortButtonsInRange(RangeofButtons, True)
7             If N > 0 Then
8                 If N = 1 Then
9                     UndoText = "Undo Remove Sort Button"
10                Else
11                    UndoText = "Undo Remove " + Format$(N, "###,###") & " Sort Buttons from " & _
                          AddressND(RangeofButtons)
12                End If
13                Application.OnUndo UndoText, "UndoRemoveSortButtons"
14            End If
15        End If

16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#RemoveSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UndoRemoveSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Implements Undo for the RemoveSortButtons method
' -----------------------------------------------------------------------------------------------------------------------
Sub UndoRemoveSortButtons()
          Dim b As Button
          Dim i As Long
          Dim SheetName As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler

2         On Error Resume Next
3         SheetName = m_WsUndoRemove.Name
4         On Error GoTo ErrHandler

5         If SheetName = vbNullString Then Exit Sub
6         If Not IsArray(m_UndoRemoveList) Then Exit Sub

7         Set SPH = CreateSheetProtectionHandler(m_WsUndoRemove)
8         Application.ScreenUpdating = False
9         For i = 1 To sNRows(m_UndoRemoveList)
10            AddSortButton b, m_WsUndoRemove.Range(m_UndoRemoveList(i, 1)), 1
11            b.OnAction = m_UndoRemoveList(i, 3)
12            SortButtonSetDirection b, CLng(m_UndoRemoveList(i, 2))
13        Next i
14        m_UndoAddList = Empty

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#UndoRemoveSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveSortButtonsInRange
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Removes sort buttons whose .TopLeftCell intersects RangeOfButtons.
'             Returns the number of buttons removed
' -----------------------------------------------------------------------------------------------------------------------
Function RemoveSortButtonsInRange(RangeofButtons As Range, Optional RecordForUndo As Boolean = False)
          Dim b As Button
          Dim N As Long
          Dim STK As clsStacker
          Dim ThisLine As Variant
1         On Error GoTo ErrHandler
2         If RecordForUndo Then
3             Set m_WsUndoRemove = ActiveSheet
4             m_UndoRemoveList = Empty
5         End If
6         ThisLine = sReshape(0, 1, 3)

7         Set STK = CreateStacker()
8         For Each b In RangeofButtons.Parent.Buttons
9             If IsSortButton(b) Then
10                If Not Application.Intersect(CellBeneathButton(b, False), RangeofButtons) Is Nothing Then
11                    If RecordForUndo Then
12                        ThisLine(1, 1) = CellBeneathButton(b, False).address
13                        ThisLine(1, 2) = SortButtonGetDirection(b)
14                        ThisLine(1, 3) = b.OnAction
15                        STK.Stack2D ThisLine
16                    End If
17                    b.Delete
18                    N = N + 1
19                End If
20            End If
21        Next b

22        If RecordForUndo Then
23            m_UndoRemoveList = STK.Report
24        End If
25        RemoveSortButtonsInRange = N

26        Exit Function
ErrHandler:
27        Throw "#RemoveSortButtonsInRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddSortButton
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Adds a single sortbutton aligned to the cell c
' -----------------------------------------------------------------------------------------------------------------------
Sub AddSortButton(ByRef b As Button, c As Range, NumHeaderRows As Long)

1         On Error GoTo ErrHandler

2         Set b = c.Parent.Buttons.Add(c.Left, c.Top, c.Width, c.Height)
3         If NumHeaderRows = 1 Then
4             b.OnAction = "SAISortButtonOnAction"
5         Else
              'Trick from http://www.tushar-mehta.com/excel/vba/xl%20objects%20and%20procedures%20with%20arguments.htm
6             b.OnAction = "'SAISortButtonOnAction " + CStr(NumHeaderRows) + " '"
7         End If

8         SortButtonSetDirection b, sbDirectionFlat

9         Exit Sub
ErrHandler:
10        Throw "#AddSortButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortButtonSetDirection
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Sets the text on a sort button. Needs Wingdings and Wingdings 3 to be installed
'             for nice-looking symbols on buttons. If they're not installed we do as best we can
'             with "-", "v" and "^" characters.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SortButtonSetDirection(b As Button, Direction As EnmsbDirection)

1         On Error GoTo ErrHandler

          Static WingDings3IsInstalled As Variant
          Static WingDingsIsInstalled As Variant

2         If IsEmpty(WingDingsIsInstalled) Then
3             WingDingsIsInstalled = sFontIsInstalled("Wingdings")
4         End If
5         If IsEmpty(WingDings3IsInstalled) Then
6             WingDings3IsInstalled = sFontIsInstalled("Wingdings 3")
7         End If

8         b.Font.Bold = False

9         Select Case Direction

              Case sbDirectionFlat
10                If WingDingsIsInstalled Then
11                    b.text = Chr$(108)
12                    b.Font.Name = "Wingdings"
13                Else
14                    b.text = "-"
15                    b.Font.Name = "Arial"
16                End If
17                b.Font.Size = 10
18                b.Font.ColorIndex = 48        'grey
19            Case sbDirectionDown
20                m_LastButtonWithArrow = b.Name
21                If WingDings3IsInstalled Then
22                    b.text = Chr$(112)
23                    b.Font.Name = "Wingdings 3"
24                Else
25                    b.text = "v"
26                    b.Font.Name = "Arial"
27                    b.Font.Bold = True
28                End If

29                b.Font.Size = 10
30                b.Font.ColorIndex = 25
31            Case sbDirectionUp
32                m_LastButtonWithArrow = b.Name
33                If WingDings3IsInstalled Then
34                    b.text = Chr$(113)
35                    b.Font.Name = "Wingdings 3"
36                Else
37                    b.text = "^"
38                    b.Font.Name = "Arial"
39                    b.Font.Bold = True
40                End If
41                b.Font.Size = 10
42                b.Font.ColorIndex = 25
43        End Select

44        Exit Sub
ErrHandler:
45        Throw "#SortButtonSetDirection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortButtonGetDirection
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Obverse of SortButtonSetDirection
' -----------------------------------------------------------------------------------------------------------------------
Private Function SortButtonGetDirection(b As Button) As EnmsbDirection
1         On Error GoTo ErrHandler

2         Select Case Asc(b.text)
              Case 108, Asc("-")
3                 SortButtonGetDirection = sbDirectionFlat
4             Case 112, Asc("v")
5                 SortButtonGetDirection = sbDirectionDown
6             Case 113, Asc("^")
7                 SortButtonGetDirection = sbDirectionUp
8             Case Else
9                 SortButtonGetDirection = sbDirectionFlat
10        End Select

11        Exit Function
ErrHandler:
12        Throw "#SortButtonGetDirection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortButtonGetSortableRange
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Given a button, what is the range to be sorted. Returns range without headers.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SortButtonGetSortableRange(b As Button, NumHeaderRows) As Range
          Dim HeadersAndFirstRow As Range
          Dim Result As Range
          Dim lo As ListObject
          Dim loTopDataRow As Range

1         On Error GoTo ErrHandler

          'Comon use case is to have sort buttons atop tables, aka ListObjects
2         For Each lo In b.Parent.ListObjects
3             Set loTopDataRow = lo.DataBodyRange.Rows(1)
4             If Not Application.Intersect(CellBeneathButton(b, True).Offset(NumHeaderRows + 1), loTopDataRow) Is Nothing Then
5                 Set SortButtonGetSortableRange = lo.DataBodyRange
6                 Exit Function
7             End If
8         Next
           
9         If Result Is Nothing Then
              'Sorting a range, not a table
10            Set HeadersAndFirstRow = RangeContainingAdjacentSortButtons(b).Offset(1).Resize(NumHeaderRows + 1)
11            Set Result = sExpand(HeadersAndFirstRow, True, True, False, True)
12            With Result
13                Set Result = .Offset(NumHeaderRows).Resize(.Rows.Count - NumHeaderRows)
14            End With
15        End If

          'Shameful hack for Portfolio sheet of SCRiPT...
16        With Result
17            If .Rows.Count > 1 Then
18                If CStr(.Cells(.Rows.Count, 2).Value) = "<Doubleclick to add trade>" Then
19                    Set Result = .Resize(.Rows.Count - 1)
20                End If
21            End If
22        End With

23        Set SortButtonGetSortableRange = Result
24        Exit Function
ErrHandler:
25        Throw "#SortButtonGetSortableRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestSortButtonGetSortableRange
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Ad hoc test harness
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestSortButtonGetSortableRange()
1         Application.GoTo SortButtonGetSortableRange(ActiveSheet.Buttons("Button 19"), 1)
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SAISortButtonOnAction
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : On Action macro when we set up sort buttons with a header row
' -----------------------------------------------------------------------------------------------------------------------
Sub SAISortButtonOnAction(Optional NumHeaders As Long = 1)
1         On Error GoTo ErrHandler
          'We know we are at the top of the call stack
2         Application.Cursor = xlDefault
3         SortButtonOnActionCore NumHeaders
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#SAISortButtonOnAction (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortButtonOnAction
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : This method is assigned to the sort buttons
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SortButtonOnActionCore(NumHeaderRows As Long)
          Dim ClickedButton As Button
          Dim Key1 As Long
          Dim order1 As XlSortOrder
          Dim SortableRange As Range
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler

2         Set SPH = CreateSheetProtectionHandler(ActiveSheet)
3         Application.ScreenUpdating = False
4         If VarType(Application.Caller) = vbString Then
5             Set ClickedButton = ActiveSheet.Buttons(Application.Caller)
6         ElseIf sElapsedTime() - LastAltBacktickTime < 0.5 Then
7             Set ClickedButton = LastAltBacktickButton
8         End If
9         Set SortableRange = SortButtonGetSortableRange(ClickedButton, NumHeaderRows)
10        If SortableRange Is Nothing Then Exit Sub        ' I think this will never be the case
11        If SortableRange.Rows.Count = 1 Then
12            ResetSortButtons RangeContainingAdjacentSortButtons(ClickedButton), False, True
13            Exit Sub
14        End If

15        CheckRangeForSorting SortableRange

          'Backup for undo functionality, button states are recorded in call to ResetSortButtons
16        Set m_wsUndoSort = ActiveSheet

          Const AllowUndo = False        'PGS 19-Nov-15. BackupRange is annoyingly slow (new & bad behaviour?) can we speed it up?

17        If AllowUndo Then
18            BackUpRange SortableRange, shUndo
19        End If

20        Key1 = CellBeneathButton(ClickedButton, False).Column - SortableRange.Column + 1
21        If SortButtonGetDirection(ClickedButton) = sbDirectionDown Then
22            order1 = xlDescending
23        Else
24            order1 = xlAscending
25        End If
26        ActiveSheet.Sort.SortFields.Clear
27        ActiveSheet.Sort.SortFields.Add Key:=SortableRange.Columns(Key1), SortOn:=xlSortOnValues, _
              order:=order1, DataOption:=xlSortNormal
28        With ActiveSheet.Sort
29            .SetRange SortableRange
30            .header = xlNo
31            .MatchCase = False
32            .Orientation = xlTopToBottom
              Dim ErrNum
33            On Error Resume Next
34            .Apply
35            ErrNum = Err.Number
36            On Error GoTo ErrHandler
37            If ErrNum <> 0 Then Throw "Range " + AddressND(SortableRange) + " could not be sorted.", True
38        End With

          Dim DoAutoFit As Variant
39        DoAutoFit = SortableRange.WrapText
40        If VarType(DoAutoFit) <> vbBoolean Then DoAutoFit = True
41        If DoAutoFit Then SortableRange.Rows.AutoFit

          'Make changes to the buttons after the sort has succeeded. It looks strange if the sort _
           fails (e.g. array formulas present) but nevertheless we changed the icon on the button
42        If order1 = xlAscending Then
43            ResetSortButtons SortableRange.Rows(-NumHeaderRows), True, True
44            SortButtonSetDirection ClickedButton, sbDirectionDown
45        Else
46            ResetSortButtons SortableRange.Rows(-NumHeaderRows), True, True
47            SortButtonSetDirection ClickedButton, sbDirectionUp
48        End If

49        If AllowUndo Then
50            If IsUndoAvailable(shUndo) Then
                  'The UndoBuffer won't exist when too many cells were sorted see constant MAX_CELLS_FOR_UNDO
51                Application.OnUndo "Undo Sort of " & AddressND(SortableRange), "SortButtonUndoSort"
52            End If
53        End If

54        Exit Sub
ErrHandler:
55        Throw "#SortButtonOnActionCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckRangeForSorting
' Author    : Philip Swannell
' Date      : 19-Mar-2016
' Purpose   : Sorting an a range of cells may fail because of array formulas - error would
'             be "You can't change part of an array". Sorting works on a range containing
'             formulas that point to cells outside the range but it's all too likely to mess
'             up the contents of the range, so this method throws an error in that circumstance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CheckRangeForSorting(TheRange As Range)
          Dim AnyArrays As Boolean
          Dim AnyFormulas As Boolean
          Dim BadCells As Range
          Dim c As Range
          Dim D As Range
          Dim DPs As Range
          Dim HashForRange As String
          Dim IsTable As Boolean
          Dim lo As ListObject
          Dim OtherRows As Range
          Dim RngFormulas As Range
          Dim ThisRow As Range
          Dim XSH As clsExcelStateHandler
          
          Static LastHashForRange As String
          Static LastRunTime As Double

1         On Error GoTo ErrHandler

          'This method can be slow when lots of cells in the range to be sorted have formulas,
          'so exit early if sorting the same range within 30 seconds
          
2         HashForRange = TheRange.Parent.Parent.Name + "|" + TheRange.Parent.Name + "|" + TheRange.address
3         If HashForRange = LastHashForRange Then
4             If sElapsedTime < LastRunTime + 30 Then
5                 LastRunTime = sElapsedTime
6                 Exit Sub
7             End If
8         End If
9         LastHashForRange = vbNullString

10        For Each lo In TheRange.Parent.ListObjects
11            If lo.DataBodyRange.address = TheRange.address Then
12                IsTable = True
13            End If
14        Next

15        On Error Resume Next
16        Set RngFormulas = TheRange.SpecialCells(xlCellTypeFormulas)
17        On Error GoTo ErrHandler

18        AnyFormulas = Not RngFormulas Is Nothing ' 100x faster than TheRange.HasFormula

19        If Not AnyFormulas Then
20            AnyArrays = False
21        ElseIf IsTable Then
22            AnyArrays = False ' because Tables (ListObjects) cannot contain array formulas
23        Else
24            AnyArrays = Not sEquals(False, TheRange.HasArray)
25        End If

26        If AnyFormulas Then
27            Set RngFormulas = TheRange.SpecialCells(xlCellTypeFormulas)

28            If AnyArrays Then
29                For Each c In RngFormulas.Cells
30                    If c.HasArray Then
31                        If c.CurrentArray.Rows.Count > 1 Then
32                            c.CurrentArray.Select
33                            Throw "The range cannot be sorted because of the array formula at cells " + AddressND(c.CurrentArray), True
34                        End If
35                    End If
36                Next c
37            End If

38            If ExcelSupportsSpill() Then
39                For Each c In RngFormulas.Cells
40                    Set DPs = Nothing
41                    On Error Resume Next
42                    Set DPs = c.SpillParent.SpillingToRange
43                    On Error GoTo ErrHandler
44                    If Not DPs Is Nothing Then
45                        If DPs.Rows.Count > 1 Then
46                            DPs.Select
47                            Throw "The range cannot be sorted because of the array formula at cells " + AddressND(DPs), True
48                        End If
49                    End If
50                Next c
51            End If

52            Set XSH = CreateExcelStateHandler(, , False) 'Temporarily switch off events since .DirectPrecedents fires the selection change event.
              'Check for references to cells in other rows of the range to be sorted, and treat as fatal
53            If TheRange.Rows.Count > 1 Then
54                For Each c In RngFormulas.Cells
55                    Set DPs = Nothing
56                    On Error Resume Next
57                    Set DPs = c.DirectPrecedents
58                    On Error GoTo ErrHandler
59                    If Not DPs Is Nothing Then
60                        Set ThisRow = TheRange.Rows(c.row - TheRange.row + 1)
61                        Set OtherRows = IntersectWithComplement(TheRange, ThisRow)
62                        Set BadCells = Application.Intersect(DPs, OtherRows)
63                        If Not BadCells Is Nothing Then
64                            Set D = BadCells.Areas(1).Cells(1, 1)
65                            Throw "The range cannot be sorted because sorting would corrupt formulas. " + _
                                  "An example is the formula at cell " + AddressND(c) + _
                                  " which refers to cell " + AddressND(D) + _
                                  ".", True
66                        End If
67                    End If
68                Next c
69            End If

              'Check for references to cells on the sheet but outside the range being sorted and check for use of absolute row address
70            For Each c In RngFormulas.Cells
71                Set DPs = Nothing
72                On Error Resume Next
73                Set DPs = c.DirectPrecedents
74                On Error GoTo ErrHandler
75                If Not DPs Is Nothing Then

76                    If Not sRangeContainsRange(TheRange, DPs) Then
77                        For Each D In IntersectWithComplement(DPs, TheRange).Cells
78                            If InStr(c.Formula, AddressND(D)) > 0 Then
                                  'More sophisticated check, using regular expression with word boundaries and also _
                                   removing literal strings from the formula so are not foxed by a weird formula such as =$A$1&"A1"
79                                If sIsRegMatch("\b" & AddressND(D) & "\b", StripLiterals(c.Formula)) Then
80                                    Application.Union(DPs, c).Select
81                                    Throw "The range cannot be sorted because sorting would corrupt formulas. " + _
                                          "An example is the formula at cell " + AddressND(c) + _
                                          " which refers to cell " + AddressND(D) + _
                                          "." + vbLf + "Suggestion: Change the formula to use an absolute row address i.e. " + Mid$(D.address, 2) + ".", True
82                                End If
83                            End If
84                        Next D
85                    End If
86                End If
87            Next c
88        End If
89        LastHashForRange = HashForRange
90        LastRunTime = sElapsedTime

91        Exit Sub
ErrHandler:
92        Throw "#CheckRangeForSorting (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StripLiterals
' Author    : Philip Swannell
' Date      : 12-Jul-2017
' Purpose   : Amends a syntactically correct Excel formula to replace all literal strings with space characters
'            e.g.  =A1+"A1"   --> =A1+"  "
' -----------------------------------------------------------------------------------------------------------------------
Private Function StripLiterals(FormulaText As String)
          Dim DQCount As Long
          Dim FormulaText2 As String
          Dim i As Long
          Const DQ = """"

1         On Error GoTo ErrHandler
2         If InStr(FormulaText, DQ) = 0 Then
3             StripLiterals = FormulaText
4         Else
5             FormulaText2 = FormulaText
6             For i = 1 To Len(FormulaText)
7                 If Mid$(FormulaText, i, 1) = DQ Then
8                     DQCount = DQCount + 1
9                 ElseIf (DQCount Mod 2 = 1) Then
10                    Mid$(FormulaText2, i, 1) = " "
11                End If
12            Next i
13            StripLiterals = FormulaText2
14        End If
15        Exit Function
ErrHandler:
16        Throw "#StripLiterals (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortButtonUndoSort
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Implements undoing of the most recent sort via a SortButton, resets button states
'             and calls modUndo.RestoreRange to put the sorted cells back to their prior state.
' -----------------------------------------------------------------------------------------------------------------------
Sub SortButtonUndoSort()
          Dim b As Button
          Dim SheetName As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

          'Check that the sheet exists, though I don't think that this method can get called when it doesn't
1         On Error GoTo ErrHandler

2         On Error Resume Next
3         SheetName = m_wsUndoSort.Name
4         On Error GoTo ErrHandler
5         If SheetName = vbNullString Then Exit Sub

6         Set SUH = CreateScreenUpdateHandler()
7         Set SPH = CreateSheetProtectionHandler(m_wsUndoSort)
8         If Not IsEmpty(m_ButtonStates) Then
9             ResetSortButtons m_wsUndoSort.Range(m_ButtonStates(1, 1)), False, False

10            If IsInCollection(m_wsUndoSort.Buttons, CStr(m_ButtonStates(2, 1))) Then
11                Set b = m_wsUndoSort.Buttons(m_ButtonStates(2, 1))
12                SortButtonSetDirection b, CLng(m_ButtonStates(3, 1))
13            End If
14        End If

15        RestoreRange

16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#SortButtonUndoSort (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResetSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : For all sort buttons in a given range, sets their caption to display as a light grey circle
' -----------------------------------------------------------------------------------------------------------------------
Sub ResetSortButtons(RangeofButtons As Range, RecordState As Boolean, BeLazy As Boolean)
          Dim b As Button
          Dim c As Range
          Dim Direction As EnmsbDirection
1         On Error GoTo ErrHandler

2         If RecordState Then
3             m_ButtonStates = Empty
4         End If

5         If BeLazy Then
6             If IsInCollection(ActiveSheet.Buttons, m_LastButtonWithArrow) Then
7                 Set b = ActiveSheet.Buttons(m_LastButtonWithArrow)
8                 Direction = SortButtonGetDirection(b)
9                 If Direction <> sbDirectionFlat Then
10                    If Not Application.Intersect(CellBeneathButton(b, False), RangeofButtons) Is Nothing Then
11                        SortButtonSetDirection b, sbDirectionFlat
12                        If RecordState Then
13                            m_ButtonStates = sArrayStack(RangeofButtons.address, b.Name, Direction)
14                        End If
15                        Exit Sub
16                    End If
17                End If
18            End If
19        End If

20        For Each b In RangeofButtons.Parent.Buttons
21            If IsSortButton(b) Then
22                Set c = CellBeneathButton(b, True)
23                If c.row = RangeofButtons.row Then
24                    If Not Application.Intersect(c, RangeofButtons) Is Nothing Then
25                        Direction = SortButtonGetDirection(b)
26                        If Direction <> sbDirectionFlat Then
27                            SortButtonSetDirection b, sbDirectionFlat
28                            If RecordState Then
29                                m_ButtonStates = sArrayStack(RangeofButtons.address, b.Name, Direction)
30                            End If

31                        End If

32                    End If
33                End If
34            End If
35        Next b

36        Exit Sub
ErrHandler:
37        Throw "#ResetSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeContainingAdjacentSortButtons
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Sort Buttons are arranged in a row of contiguous cells, given a sort button
'             this method returns the range that contains all its adjacent sort buttons.
' -----------------------------------------------------------------------------------------------------------------------
Private Function RangeContainingAdjacentSortButtons(SortButton As Button) As Range
          Dim AllColNos As Variant
          Dim b As Button
          Dim c As Range
          Dim ColNo As Long
          Dim i As Long
          Dim RowNo As Long
          Dim STK As clsStacker
          Dim ws As Worksheet
1         On Error GoTo ErrHandler

2         Set ws = SortButton.TopLeftCell.Parent

          Dim UseCachedResult As Boolean
          Static ButtonCount As Long
          Static LastCallTime As Double
          Static ResultRange As Range
          Static TheSheet As Worksheet

          'Unfortunately this method can be slow when there are a large number of buttons on the sheet _
           i.e. 0.4 seconds for 500 buttons on sheet. So we cache the result and use that previously _
           calculated result if possible. The caching is not "Bomb Proof" - consider what happens if _
           the user cuts and pastes one of the sort buttons to a different location. For that reason _
           we don't trust cached results more than 10 seconds old
3         On Error Resume Next
4         UseCachedResult = (Now - LastCallTime < (10 / 24 / 60 / 60)) And _
              (ws Is TheSheet) And (ws.Buttons.Count = ButtonCount) And _
              ButtonCount > 499 And _
              Not (Application.Intersect(CellBeneathButton(SortButton, False), ResultRange) Is Nothing)

5         On Error GoTo ErrHandler

6         If UseCachedResult Then
7             Set RangeContainingAdjacentSortButtons = ResultRange
8         Else
9             With CellBeneathButton(SortButton, False)
10                RowNo = .row
11                ColNo = .Column
12            End With

13            Set STK = CreateStacker()
14            For Each b In ws.Buttons
15                If IsSortButton(b) Then
16                    Set c = CellBeneathButton(b, True)
17                    If c.row = RowNo Then
18                        STK.Stack0D c.Column
19                    End If
20                End If
21            Next b

22            AllColNos = STK.Report
23            AllColNos = sSortedArray(AllColNos)

              Dim LeftCell As Range
              Dim LeftPos As Long
              Dim MatchID As Long
              Dim RightCell As Range
              Dim RightPos As Long
24            MatchID = Application.WorksheetFunction.Match(ColNo, AllColNos, 0)
25            LeftPos = MatchID
26            RightPos = MatchID

27            For i = MatchID To 2 Step -1
28                If AllColNos(LeftPos - 1, 1) >= AllColNos(LeftPos, 1) - 1 Then
29                    LeftPos = LeftPos - 1
30                Else
31                    Exit For
32                End If
33            Next i

34            For i = MatchID To sNRows(AllColNos) - 1
35                If AllColNos(RightPos + 1, 1) <= AllColNos(RightPos, 1) + 1 Then
36                    RightPos = RightPos + 1
37                Else
38                    Exit For
39                End If
40            Next i

41            Set LeftCell = ws.Cells(RowNo, AllColNos(LeftPos, 1))
42            Set RightCell = ws.Cells(RowNo, AllColNos(RightPos, 1))

              'Set static variables for speed of call next time
43            Set ResultRange = ws.Range(LeftCell, RightCell)
44            Set TheSheet = ws
45            ButtonCount = ws.Buttons.Count
46            LastCallTime = Now()
47        End If

48        Set RangeContainingAdjacentSortButtons = ResultRange

49        Exit Function
ErrHandler:
50        Throw "#RangeContainingAdjacentSortButtons (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Test_RangeContainingAdjacentSortButtons
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Ad hoc test harness
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Test_RangeContainingAdjacentSortButtons()
1         Application.GoTo RangeContainingAdjacentSortButtons(ActiveSheet.Buttons("Button 35"))
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestSortButton
' Author    : Philip Swannell
' Date      : 16-Oct-2013
' Purpose   : Speed bench test
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestSortButton()
          Dim i As Long
          Dim R As Range
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double
1         For i = 100 To 500 Step 100
2             Application.GoTo ActiveSheet.Cells(1, 1).Resize(1, i)
3             t1 = sElapsedTime()
4             AddSortButtons , 1
5             t2 = sElapsedTime()
6             Set R = RangeContainingAdjacentSortButtons(ActiveSheet.Buttons(1))
7             t3 = sElapsedTime()
8             RemoveSortButtons
9             t4 = sElapsedTime
10            Debug.Print i, "Add", t2 - t1, "RCASB", t3 - t2, "Remove", t4 - t3
11        Next i
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CellBeneathButton
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : When a button is positioned above a cell, TopLeftCell has a nasty habit of
'             returning the "wrong" cell and this is particularly the case in Office 2010 when there are
'             zero width columns nearby. This method returns the cell whose centre point is
'             closest (or equal closest?) to the centre point of the button
' -----------------------------------------------------------------------------------------------------------------------
Private Function CellBeneathButton(b As Button, SnapToGrid As Boolean)
          Dim TheCell As Range

1         On Error GoTo ErrHandler

          'First check for perfect alignment
2         With b
3             Set TheCell = .TopLeftCell
4             If .Top = TheCell.Top Then
5                 If .Left = TheCell.Left Then
6                     If .Width = TheCell.Width Then
7                         If .Height = TheCell.Height Then
8                             Set CellBeneathButton = TheCell
9                             Exit Function
10                        End If
11                    End If
12                End If
13            End If
14        End With

          'PGS 6 Jan 2020. Check for very strange errors seen when using two screens of high and different resolution - error only seen when worksheet displayed on one of the two screens.
          'TODO not yet recreated the "bad" setup to test if this code does catch it!
15        With b
16            If TheCell.Width > 0 And TheCell.Height > 0 Then
17                If .Top > TheCell.Top + TheCell.Height Or _
                      (.Top + .Height) < TheCell.Top Then
18                    Throw "Assertion failed for button '" + b.Name + "' on worksheet '" + TheCell.Parent.Name + "'. It's position (Top, Left, Height and Width properties) are inconsistent with the position of its TopLeftCell. This is a possible bug in Excel, and the " + gAddinName + " 'Sort Buttons' don't work correctly when that bug is manifested"
19                End If
20            End If
21        End With

          'Button is slightly out of alignment?
22        Do While AdjacentIsCloser(TheCell, b, xlToLeft)
23            Set TheCell = TheCell.Offset(, -1)
24        Loop
25        Do While AdjacentIsCloser(TheCell, b, xlToRight)
26            Set TheCell = TheCell.Offset(, 1)
27        Loop
28        Do While AdjacentIsCloser(TheCell, b, xlUp)
29            Set TheCell = TheCell.Offset(-1)
30        Loop
31        Do While AdjacentIsCloser(TheCell, b, xlDown)
32            Set TheCell = TheCell.Offset(1)
33        Loop

34        If TheCell.Width = 0 Or TheCell.Height = 0 Then
              'because which cell is underneath the button is not defined, thanks to zero width or zero height cells, so we rely on the .TopLeftCell property
35            Set CellBeneathButton = b.TopLeftCell
36        Else
37            If SnapToGrid Then
38                With b
39                    .Top = TheCell.Top
40                    .Left = TheCell.Left
41                    .Height = TheCell.Height
42                    .Width = TheCell.Width
43                    .Placement = xlMoveAndSize
44                End With
45            End If
46            Set CellBeneathButton = TheCell
47        End If

48        Exit Function
ErrHandler:
49        Throw "#CellBeneathButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AdjacentIsCloser
' Author    : Philip Swannell
' Date      : 09-Nov-2016
' Purpose   : Is the centre of the cell adjacent to C in the specified direction closer (or at least as close)
'             to the centre of the button than is the centre of c
' -----------------------------------------------------------------------------------------------------------------------
Private Function AdjacentIsCloser(c As Range, b As Button, Direction As XlDirection)

          Dim bx As Double
          Dim by As Double
          Dim cx As Double
          Dim cy As Double
1         On Error GoTo ErrHandler
2         bx = b.Left + b.Width / 2
3         by = b.Top + b.Height / 2
4         cx = c.Left + c.Width / 2
5         cy = c.Top + c.Height / 2

6         Select Case Direction
              Case xlDown
7                 If c.row = c.Parent.Rows.CountLarge Then
8                     AdjacentIsCloser = False
9                 Else
10                    AdjacentIsCloser = Abs((c.Offset(1).Top + c.Offset(1).Height / 2) - by) <= Abs(cy - by)
11                End If
12            Case xlUp
13                If c.row = 1 Then
14                    AdjacentIsCloser = False
15                Else
16                    AdjacentIsCloser = Abs((c.Offset(-1).Top + c.Offset(-1).Height / 2) - by) <= Abs(cy - by)
17                End If
18            Case xlToLeft
19                If c.Column = 1 Then
20                    AdjacentIsCloser = False
21                Else
22                    AdjacentIsCloser = Abs((c.Offset(0, -1).Left + c.Offset(0, -1).Width / 2) - bx) <= Abs(cx - bx)
23                End If
24            Case xlToRight
25                If c.Column = c.Parent.Columns.Count Then
26                    AdjacentIsCloser = False
27                Else
28                    AdjacentIsCloser = Abs((c.Offset(0, 1).Left + c.Offset(0, 1).Width / 2) - bx) <= Abs(cx - bx)
29                End If
30        End Select
31        Exit Function
ErrHandler:
32        Throw "#AdjacentIsCloser (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
