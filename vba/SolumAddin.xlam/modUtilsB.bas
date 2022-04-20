Attribute VB_Name = "modUtilsB"
'Module that contains code called from the Ribbon.
Option Explicit

Private Declare PtrSafe Function FindWindowExA Lib "USER32" ( _
    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr

Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hWnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
    
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long

Sub wefsfe()
          Dim PID As Long
1         GetWindowThreadProcessId Application.hWnd, PID
2         Debug.Print PID
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MyPaste
' Author     : Philip Swannell
' Date       : 04-Apr-2019
' Purpose    : Pastes an array to a range while
'              a) doing no conversion of string to non-string; and
'              b) preserving horizontal alignment
' -----------------------------------------------------------------------------------------------------------------------
Sub MyPaste(R As Range, ByVal v As Variant)
          Dim CopyOfErr As String
          Dim OldTNK As Boolean

1         On Error GoTo ErrHandler
2         OldTNK = Application.TransitionNavigKeys
3         Force2DArrayR v
4         v = sArrayExcelString(v)
5         If OldTNK Then Application.TransitionNavigKeys = False
6         R.Value2 = v
7         If OldTNK Then Application.TransitionNavigKeys = True

8         Exit Sub
ErrHandler:
9         CopyOfErr = "#MyPaste (line " & CStr(Erl) + "): " & Err.Description & "!"
10        If OldTNK Then Application.TransitionNavigKeys = True
11        Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InsertFolderName
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Paste the name of a folder to the active cell. Available from the Ribbon.
'             and assigned to Ctrl+Alt+Shift+F
' -----------------------------------------------------------------------------------------------------------------------
Sub InsertFolderName()
          Dim ChosenFolder As Variant
1         On Error GoTo ErrHandler
2         If ActiveCell Is Nothing Then Exit Sub

3         ChosenFolder = FolderPicker(CStr(ActiveCell.Value), "Paste", "Paste Folder Name to Cell: Select Folder", "InsertFolderName", True, ActiveCell)
4         If VarType(ChosenFolder) = vbString Then
5             If Not UnprotectAsk(ActiveSheet, , ActiveCell) Then Exit Sub
6             BackUpRange ActiveCell, shUndo
7             ActiveCell.Value = "'" + ChosenFolder
8             Application.OnUndo "Undo Paste Folder Name" & " to " & AddressND(ActiveCell), "RestoreRange"
9         End If
10        Application.OnRepeat "Repeat Insert Folder Name", "InsertFolderName"

11        Exit Sub
ErrHandler:
12        SomethingWentWrong "#InsertFolderName (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, "Paste Folder Name"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddNewWorkbook
' Author    : Philip Swannell
' Date      : 27-Jun-2013
' Purpose   : Adds a new workbook - assigned to ctrl+N
' -----------------------------------------------------------------------------------------------------------------------
Sub AddNewWorkbook()
1         Application.Workbooks.Add
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CtrlF6Response
' Author    : Philip Swannell
' Date      : 17-Jun-2014
' Purpose   : Excel's response to Ctrl+F6 appears to have changed in Office 2013. See discussion at
'             http://answers.microsoft.com/en-us/office/forum/office_2013_release-excel/cycling-between-open-windows-using-ctrlf6-or/7314b632-7f48-44f1-8e96-7d7000ac86d8
'             and ActiveWindow.ActivateNext has the same not-very-desirable behaviour.
'             This method switches to the next visible window, having sorted Windows alphabetically by their caption
'             PGS 7 Apr 2016 - Excel 2016 has same undesirable feature.
' -----------------------------------------------------------------------------------------------------------------------
Sub CtrlF6Response()
          Dim i As Long
          Dim NumVisibleWindows As Long
          Dim TheWindows()
          Dim w As Window

1         On Error GoTo ErrHandler
2         For Each w In Application.Windows
3             If w.Visible Then
4                 NumVisibleWindows = NumVisibleWindows + 1
5             End If
6         Next w

7         If NumVisibleWindows <= 1 Then Exit Sub

8         ReDim TheWindows(1 To NumVisibleWindows)
9         For Each w In Application.Windows
10            If w.Visible Then
11                i = i + 1
12                TheWindows(i) = w.caption
13            End If
14        Next w

15        If ActiveWindow Is Nothing Then        'Happens when user has opened a workbook in "Protected View" _
                                                  e.g. opened a workbook they've just downloaded from the web, and it does not _
                                                  seem to be possible to write VBA code that switches focus to another sheet - _
                                                  I tried the approaches commented out below.
              'Application.Windows(TheWindows(1)).Activate
              'Application.Workbooks(1).Activate
              'Application.SendKeys "+^{F6}"
16            Exit Sub
17        End If

18        QuickSort TheWindows, LBound(TheWindows, 1), UBound(TheWindows, 1)
19        For i = 1 To NumVisibleWindows
20            If ActiveWindow.caption = TheWindows(i) Then
21                Set w = Application.Windows(TheWindows(i Mod NumVisibleWindows + 1))
22                w.Activate
23                If w.WindowState = xlMinimized Then
24                    w.WindowState = xlNormal
25                End If
26                Exit Sub
27            End If
28        Next

29        Exit Sub
ErrHandler:
30        SomethingWentWrong "#CtrlF6Response (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, "CtrlF6Response"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSort
' Author    : Copied by Philip Swannell from http://stackoverflow.com/questions/152319/vba-array-sort-function
' Date      : 07-Jan-2015
' Purpose   : Sort a 1 dimensional Array using QuickSort algorithm and ordering inherited
'             from VBA < operator
' -----------------------------------------------------------------------------------------------------------------------
Private Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

          Dim pivot As Variant
          Dim tmpHi As Long
          Dim tmpLow As Long
          Dim tmpSwap As Variant

1         On Error GoTo ErrHandler
2         tmpLow = inLow
3         tmpHi = inHi

4         pivot = vArray((inLow + inHi) \ 2)

5         Do While (tmpLow <= tmpHi)

6             Do While (vArray(tmpLow) < pivot And tmpLow < inHi)
7                 tmpLow = tmpLow + 1
8             Loop

9             Do While (pivot < vArray(tmpHi) And tmpHi > inLow)
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

20        If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
21        If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
22        Exit Sub
ErrHandler:
23        Throw "#QuickSort (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CalcSelection
' Author    : Philip Swannell
' Date      : 13-May-2015
' Purpose   : At Thu Uyen's request. Attached to Ctrl+Shift+E
' -----------------------------------------------------------------------------------------------------------------------
Sub CalcSelection()
          Dim DidPivotTables As Boolean
          Dim ExpandedRange As Range
          Dim Message As String
          Dim p As PivotTable
          Dim SafeCalcRes As Boolean
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double
1         On Error GoTo ErrHandler
2         EnsureAppObjectExists
3         If Not Selection Is Nothing Then
4             If TypeName(Selection) = "Range" Then
5                 If ExcelSupportsSpill() Then
6                     SafeCalcRes = False
7                 Else
8                     t1 = sElapsedTime()
9                     SafeCalcRes = SafeCalc(Selection)
10                    t2 = sElapsedTime()
11                End If
12                If Not SafeCalcRes Then
13                    If SheetIsProtectedWithPassword(Selection.Parent) Then
14                        Throw "You cannot use that command when both:" + vbLf + "a) the sheet is protected with a password" + vbLf + "b) there is an array formula that is only partly inside the current selection." + vbLf + _
                              "You must either unprotect the sheet or or ensure that there are no array formulas that are part inside and part outside the selection."
15                    End If

16                    Set ExpandedRange = ExpandRangeToIncludeEntireArrayFormulas(Selection)
17                    t1 = sElapsedTime()
18                    SafeCalcRes = SafeCalc(ExpandedRange)
19                    If Not SafeCalcRes Then Throw "Unknown error in range calculation"
20                    t2 = sElapsedTime()
                      'avoid changing the viewport when we select the expanded range
                      Dim oldScrollColumn As Long
                      Dim oldScrollRow As Long
21                    oldScrollColumn = ActiveWindow.ScrollColumn
22                    oldScrollRow = ActiveWindow.ScrollRow
23                    Set SUH = CreateScreenUpdateHandler()
24                    ExpandedRange.Select
25                    ActiveWindow.ScrollColumn = oldScrollColumn
26                    ActiveWindow.ScrollRow = oldScrollRow
27                End If
28            End If
29            For Each p In ActiveSheet.PivotTables
30                t3 = sElapsedTime()
31                If Not Application.Intersect(p.TableRange2, Selection) Is Nothing Then
32                    If SPH Is Nothing Then Set SPH = CreateSheetProtectionHandler(ActiveSheet)
33                    DidPivotTables = True
34                    p.RefreshTable
35                End If
36                t4 = sElapsedTime
37            Next p
38            If DidPivotTables Then
39                Message = Format$((t2 - t1 + t4 - t3) * 1000, "0.00") + " milliseconds to calculate " + AddressND(Selection) + " on sheet '" + Selection.Parent.Name + "' in workbook '" + Selection.Parent.Parent.Name + "', including refreshing pivot tables."
40            Else
41                Message = Format$((t2 - t1) * 1000, "0.00") + " milliseconds to calculate " + AddressND(Selection) + " on sheet '" + Selection.Parent.Name + "' in workbook '" + Selection.Parent.Parent.Name + "'"
42            End If
43            TemporaryMessage Message, , vbNullString
44        End If
45        Application.OnRepeat "Repeat Calculate Selection", "CalcSelection"
46        Exit Sub
ErrHandler:
47        SomethingWentWrong "#CalcSelection (line " & CStr(Erl) + "): " & Err.Description & "!", , "CalcSelection"
48        Application.OnRepeat "Repeat Calculate Selection", "CalcSelection"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CalcActiveSheet
' Author     : Philip Swannell
' Date       : 19-Nov-2018
' Purpose    : Calculate the active sheet, with display of time taken in application status bar. Assigned to Ctrl Alt Shift C
' -----------------------------------------------------------------------------------------------------------------------
Sub CalcActiveSheet()
          Dim Message As String
          Dim t1 As Double
          Dim t2 As Double
1         On Error GoTo ErrHandler
2         If Not ActiveSheet Is Nothing Then
3             t1 = sElapsedTime()
4             ActiveSheet.Calculate
5             t2 = sElapsedTime()
6             Message = Format$((t2 - t1) * 1000, "0.00") + " milliseconds to calculate worksheet '" + ActiveSheet.Name + "' in workbook '" + ActiveSheet.Parent.Name + "'"
7             TemporaryMessage Message, , vbNullString
8             Application.OnRepeat "Repeat Calculate Active Sheet", "CalcActiveSheet"
9         End If

10        Exit Sub
ErrHandler:
11        SomethingWentWrong "#CalcActiveSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeCalc
' Author    : Philip Swannell
' Date      : 13-May-2015
' Purpose   : Sub-routine of CalcSelection
' -----------------------------------------------------------------------------------------------------------------------
Private Function SafeCalc(R As Range) As Boolean
1         On Error GoTo ErrHandler
2         R.Calculate
3         SafeCalc = True
4         Exit Function
ErrHandler:
5         SafeCalc = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ArrangeWindows
' Author     : Philip Swannell
' Date       : 26-Jun-2018
' Purpose    : Similar to Excel Ribbon View > Arrange All except that's become unreliable - sometimes works,
'              sometimes does not work.
'              Assigned to Ctrl+Alt+Shift+W, and available from Ribbon.
' -----------------------------------------------------------------------------------------------------------------------
Sub ArrangeWindows()
          Dim AllXLs As Collection
          Dim chAllExcels As String
          Dim chBothScreens As String
          Dim chCascaded As String
          Dim chLeftScreen As String
          Dim chMaximised As String
          Dim chRibbonCollaped As String
          Dim chRibbonExpanded As String
          Dim chRibbonUnchanged As String
          Dim chRightScreen As String
          Dim chSideBySide As String
          Dim chStacked As String
          Dim chThisExcel As String
          Dim chTiled As String
          Dim CopyOfErr As String
          Dim i As Long
          Dim InstanceChoices As Variant
          Dim LeftScreenHeight As Long
          Dim LeftScreenLeft As Long
          Dim LeftScreenTop As Long
          Dim LeftScreenWidth As Long
          Dim NewState
          Dim NumWindows As Long
          Dim Orientation As String
          Dim PreviousChoices As Variant
          Dim ReallyHaveTwoScreens As Boolean
          Dim RibbonHeight As String
          Dim RightScreenHeight As Long
          Dim RightScreenLeft As Long
          Dim RightScreenTop As Long
          Dim RightScreenWidth As Long
          Dim scrH As Long
          Dim scrL As Long
          Dim scrT As Long
          Dim scrW As Long
          Dim TopText As Variant
          Dim WhichScreens As String
          Dim wn As Window
          Dim wnH As Long
          Dim wnL As Long
          Dim wnT As Long
          Dim wnW As Long
          Dim xl As Application
          Const Title = "Arrange Windows (" + gAddinName + ")"

1         On Error GoTo ErrHandler

2         GetExcelInstances AllXLs

3         For Each xl In AllXLs
4             For Each wn In xl.Windows
5                 If wn.Visible Then NumWindows = NumWindows + 1
6             Next
7         Next xl

8         If NumWindows = 0 Then GoTo EarlyExit
9         chThisExcel = "This Excel Instance"
10        Select Case AllXLs.Count
              Case 0, 1
11                InstanceChoices = CreateMissing()
12            Case 2
13                chAllExcels = "Both Excel instances"
14                InstanceChoices = sArrayStack(chThisExcel, chAllExcels)
15            Case 3 To 9
16                chAllExcels = "All " & Choose(AllXLs.Count - 2, "three", "four", "five", "six", "seven", "eight", "nine") + " Excel instances"
17                InstanceChoices = sArrayStack(chThisExcel, chAllExcels)
18            Case Else
19                chAllExcels = "All " + CStr(AllXLs.Count) + " Excel instances"
20                InstanceChoices = sArrayStack(chThisExcel, chAllExcels)
21        End Select
22        TopText = "Display " + CStr(NumWindows) + " window" + IIf(NumWindows = 1, vbNullString, "s") + ":"

23        FindScreenDimensions LeftScreenWidth, LeftScreenHeight, LeftScreenTop, LeftScreenLeft, RightScreenWidth, RightScreenHeight, RightScreenTop, RightScreenLeft, ReallyHaveTwoScreens

          Dim TheChoices

24        chTiled = "Tiled"
25        chSideBySide = "Side by side"
26        chStacked = "Stacked"
27        chMaximised = "Maximised"
28        chCascaded = "Cascaded"
29        If ReallyHaveTwoScreens Then
30            TopText = sArrayRange(TopText, "on:", "with ribbons:", "windows of:")
31            chBothScreens = "Both screens"
32            chLeftScreen = "the &Left screen"
33            chRightScreen = "the &Right screen"
34        Else
35            TopText = sArrayRange(TopText, "on the:", "with ribbons:", "windows of:")
36            chBothScreens = "Whole screen"
37            chLeftScreen = "&Left half of screen"
38            chRightScreen = "&Right half of screen"
39        End If

40        chRibbonUnchanged = "Unchanged"
41        chRibbonCollaped = "Collapsed"
42        chRibbonExpanded = "Expanded"

43        TheChoices = sArrayRange(sArrayStack(chTiled, chSideBySide, chStacked, chCascaded, chMaximised), sArrayStack(chBothScreens, chLeftScreen, chRightScreen), sArrayStack(chRibbonUnchanged, chRibbonCollaped, chRibbonExpanded), InstanceChoices)

44        PreviousChoices = sArrayRange(GetSetting(gAddinName, "ArrangeWindows", "Arrangement", chTiled), _
              GetSetting(gAddinName, "ArrangeWindows", "On", chBothScreens), _
              GetSetting(gAddinName, "ArrangeWindows", "Ribbon", chRibbonUnchanged))

45        NewState = ShowOptionButtonDialog(TheChoices, Title, TopText, PreviousChoices, , False)  ', CheckBoxText, CheckBoxValue)

46        If IsEmpty(NewState) Then GoTo EarlyExit
47        Orientation = NewState(1, 1)
48        WhichScreens = NewState(1, 2)

49        SaveSetting gAddinName, "ArrangeWindows", "Arrangement", Orientation
50        SaveSetting gAddinName, "ArrangeWindows", "On", WhichScreens
51        RibbonHeight = NewState(1, 3)
52        SaveSetting gAddinName, "ArrangeWindows", "Ribbon", RibbonHeight

53        If AllXLs.Count > 1 Then
54            If NewState(1, 3) = chAllExcels Then
55                Set AllXLs = New Collection
56                AllXLs.Add Excel.Application
57                NumWindows = 0
58                For Each xl In AllXLs
59                    For Each wn In xl.Windows
60                        If wn.Visible Then NumWindows = NumWindows + 1
61                    Next
62                Next xl
63            End If
64        End If

65        Select Case WhichScreens
              Case chBothScreens
66                scrW = LeftScreenWidth + RightScreenWidth: scrH = RightScreenHeight: scrT = LeftScreenTop: scrL = LeftScreenLeft
67            Case chLeftScreen
68                scrW = LeftScreenWidth: scrH = LeftScreenHeight: scrT = LeftScreenTop: scrL = LeftScreenLeft
69            Case chRightScreen
70                scrW = RightScreenWidth: scrH = RightScreenHeight: scrT = RightScreenTop: scrL = RightScreenLeft
71        End Select

72        If Orientation = chTiled Then
              Dim Coordinates
73            Coordinates = TileCoords(IIf(ReallyHaveTwoScreens And WhichScreens = chBothScreens, 2, 1), NumWindows)
74        End If

75        i = 0
76        For Each xl In AllXLs
77            For Each wn In xl.Windows

78                If wn.Visible Then
79                    i = i + 1
80                    If Orientation = chStacked Then
81                        wnT = scrT + (i - 1) / NumWindows * scrH
82                        wnL = scrL
83                        wnW = scrW
84                        wnH = 1 / NumWindows * scrH
85                    ElseIf Orientation = chSideBySide Then
86                        wnT = scrT
87                        wnL = scrL + (i - 1) / NumWindows * scrW
88                        wnW = 1 / NumWindows * scrW
89                        wnH = scrH
90                    ElseIf Orientation = chTiled Then    'Tiled
91                        wnT = scrT + scrH * Coordinates(i, 1)
92                        wnL = scrL + scrW * Coordinates(i, 2)
93                        wnH = scrH * Coordinates(i, 3)
94                        wnW = scrW * Coordinates(i, 4)
95                    ElseIf Orientation = chMaximised Then
96                        wnT = scrT
97                        wnL = scrL
98                        wnH = scrH
99                        wnW = scrW
100                   ElseIf Orientation = chCascaded Then
                          Dim nudge As Double
101                       If NumWindows <= 1 Then
102                           nudge = 15
103                       Else
104                           nudge = SafeMin(15, SafeMin(scrH, scrW) * 0.5 / NumWindows)
105                       End If

106                       wnH = scrH - (NumWindows - 1) * nudge
107                       wnW = scrW - (NumWindows - 1) * nudge
108                       wnT = scrT + (i - 1) * nudge
109                       wnL = scrL + (i - 1) * nudge
110                       wn.Activate ' needed to get the z-order correct
111                   End If
112                   wn.WindowState = xlNormal
113                   If wn.Height <> wnH Then wn.Height = wnH
114                   If wn.Width <> wnW Then wn.Width = wnW
115                   If wn.Top <> wnT Then wn.Top = wnT
116                   If wn.Left <> wnL Then wn.Left = wnL
117                   If ((Not ReallyHaveTwoScreens) And Orientation = chMaximised And WhichScreens = chBothScreens) Or _
                          (ReallyHaveTwoScreens And Orientation = chMaximised And WhichScreens <> chBothScreens) Then
118                       wn.WindowState = xlMaximized
119                   End If

120                   If RibbonHeight = chRibbonCollaped Then
121                       wn.Activate
122                       If xl.CommandBars("Ribbon").Height > 136 Then
123                           xl.CommandBars.ExecuteMso "MinimizeRibbon"
124                       End If
125                   ElseIf RibbonHeight = chRibbonExpanded Then
126                       wn.Activate
127                       If xl.CommandBars("Ribbon").Height <= 136 Then
128                           xl.CommandBars.ExecuteMso "MinimizeRibbon"
129                       End If
130                   End If
131               End If
132           Next wn
133       Next xl

EarlyExit:
134       Application.OnRepeat "Repeat Arrange Windows", "ArrangeWindows"

135       Exit Sub
ErrHandler:
136       CopyOfErr = Err.Description
137       If Not wn Is Nothing Then CopyOfErr = CopyOfErr + " (wn.Caption = " & wn.caption & ")"
138       SomethingWentWrong "#ArrangeWindows (line " & CStr(Erl) + "): " & CopyOfErr & "!"
139       Application.OnRepeat "Repeat Arrange Windows", "ArrangeWindows"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TileColumnCounts
' Author     : Philip Swannell
' Date       : 09-Jul-2018
' Purpose    : Emulates Excel's algorithm behind Window > Arrange > Tile. Say we want to tile 5 windows on a single screen
'              returns will be an array {2;2;3} indicating that there should be three columns of tiles, the first with two
'              tiles and the next two with three tiles each.
' Parameters :
'  NumScreens: 1 or 2 - if 2 then the function recursively calls itself to "place" (approx) half tiles on one screen and
'              the other half on the second screen.
'  NumWindows: How many windows aka tiles to place.
' -----------------------------------------------------------------------------------------------------------------------
Private Function TileColumnCounts(NumScreens As Long, NumWindows As Long)
          Dim NumCols As Double
          Dim NumTilesA As Double
          Dim NumTilesB As Long
          Dim NumColsA As Long    'number of columns with the lower-by-1 number of tiles
          Dim i As Long
          Dim NumColsB As Long
          Dim Res
          Dim Result()

1         If NumWindows = 1 Then
2             TileColumnCounts = sReshape(1, 1, 1)
3             Exit Function
4         End If

5         On Error GoTo ErrHandler
6         If NumScreens = 1 Then
7             NumCols = Sqr(NumWindows)
8             If NumCols <> CLng(NumCols) Then NumCols = CLng(NumCols + 0.5)
9             NumTilesA = NumWindows / NumCols
10            If NumTilesA <> CLng(NumTilesA) Then NumTilesA = CLng(NumTilesA - 0.5)

11            NumTilesB = NumTilesA + 1
12            NumColsB = NumWindows - NumCols * NumTilesA
13            NumColsA = NumCols - NumColsB
14            ReDim Result(1 To NumCols, 1 To 1)
15            For i = 1 To NumColsA
16                Result(i, 1) = NumTilesA
17            Next
18            For i = NumColsA + 1 To NumCols
19                Result(i, 1) = NumTilesB
20            Next
21            TileColumnCounts = Result

22        ElseIf NumScreens = 2 Then
23            If NumWindows Mod 2 = 0 Then
24                Res = TileColumnCounts(1, NumWindows / 2)
25                TileColumnCounts = sArrayStack(Res, Res)
26            Else
27                TileColumnCounts = sArrayStack(TileColumnCounts(1, (NumWindows - 1) / 2), TileColumnCounts(1, (NumWindows + 1) / 2))
28            End If
29        Else
30            Throw "NumWindows must be 1 or 2"
31        End If

32        Exit Function
ErrHandler:
33        Throw "#TileColumnCounts (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TileCoords
' Author     : Philip Swannell
' Date       : 09-Jul-2018
' Purpose    : For a "tiling" across NumScreens, returns a four column array with the coordinates of each of the tiles
'              columns are Top, Left, Height Width and coordinate system is (0,0) = top left of screens, (1,1)= bottom right of screens
' -----------------------------------------------------------------------------------------------------------------------
Private Function TileCoords(NumScreens As Long, NumWindows As Long)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NumCols As Long
          Dim NumPerCol As Variant
          Dim NumRows As Long
          Dim Result As Variant
          
1         On Error GoTo ErrHandler

2         NumPerCol = TileColumnCounts(NumScreens, NumWindows)
3         NumCols = sNRows(NumPerCol)
4         For i = 1 To NumCols
5             NumPerCol(i, 1) = CLng(NumPerCol(i, 1))
6         Next i
7         If sColumnSum(NumPerCol)(1, 1) <> NumWindows Then Throw "Assertion failed"

8         Result = sReshape(0, NumWindows, 4)
9         For i = 1 To NumCols
10            NumRows = NumPerCol(i, 1)
11            For j = 1 To NumRows
12                k = k + 1
13                Result(k, 1) = (j - 1) / NumRows    'Top as proportion of screen
14                Result(k, 2) = (i - 1) / NumCols    'Left as proportion of screen
15                Result(k, 3) = 1 / NumRows    'Height as proportion of screen
16                Result(k, 4) = 1 / NumCols    'Width as proportion of screen
17            Next j
18        Next i

19        TileCoords = Result

20        Exit Function
ErrHandler:
21        Throw "#TileCoords (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ToggleWindow
' Author    : Philip Swannell
' Date      : 13-Nov-2015
' Purpose   : If PC is set up with two screens side-by-side
'             Toggles the active window between three states:
'             1) Maximised in left screen
'             2) Maximised in Right screen
'             3) Stretched across both screens
'             If PC has only one screen then the three states are:
'             1) Left half of screen
'             2) Right half of screen
'             3) Entire screen
'             Assigned to Ctrl+Shift+W
' -----------------------------------------------------------------------------------------------------------------------
Sub ToggleWindow(Optional NewState As Long)        ' NewState: 0 = Ask user, 1 = Left, 2 = Right, 3 = Both
          Dim chBoth As String
          Dim chLeft As String
          Dim chRight As String
          Dim CurrentState
          Dim LeftScreenHeight As Long
          Dim LeftScreenLeft As Long
          Dim LeftScreenTop As Long
          Dim LeftScreenWidth As Long
          Dim ReallyHaveTwoScreens As Boolean
          Dim RightScreenHeight As Long
          Dim RightScreenLeft As Long
          Dim RightScreenTop As Long
          Dim RightScreenWidth As Long
          Dim WidthShouldBe As Long
          Const Title = "Toggle Window (" + gAddinName + ")"

1         On Error GoTo ErrHandler
2         If ActiveWindow Is Nothing Then Exit Sub

3         FindScreenDimensions LeftScreenWidth, LeftScreenHeight, LeftScreenTop, LeftScreenLeft, RightScreenWidth, RightScreenHeight, RightScreenTop, RightScreenLeft, ReallyHaveTwoScreens
4         If ActiveWindow.Width > (SafeMax(LeftScreenWidth, RightScreenWidth) + LeftScreenWidth + RightScreenWidth) / 2 Then
5             CurrentState = 3
6         ElseIf ActiveWindow.Left < (LeftScreenLeft + RightScreenLeft) / 2 Then
7             CurrentState = 1
8         Else
9             CurrentState = 2
10        End If
11        If NewState = 0 Then
12            If ReallyHaveTwoScreens Then
13                chLeft = "On the &left screen"
14                chRight = "On the &right screen"
15                chBoth = "Across &both screens"
16            Else
17                chLeft = "On the &left half of the screen"
18                chRight = "On the &right half of the screen"
19                chBoth = "On the &whole screen"
20            End If
21            NewState = ShowOptionButtonDialog(sArrayStack(chLeft, chRight, chBoth), _
                  Title, "Where do you want to put " + _
                  vbLf + "'" + ActiveWindow.caption + "'?", CurrentState, , True)
22        End If

23        With ActiveWindow
24            If NewState = 2 Then
                  'Move to right window
25                .WindowState = xlNormal
26                .Width = RightScreenWidth
27                .Height = RightScreenHeight
28                .Left = RightScreenLeft
29                .Top = RightScreenTop
30                If ReallyHaveTwoScreens Then .WindowState = xlMaximized
31            ElseIf NewState = 3 Then
                  'Move to both screens.
32                .WindowState = xlNormal
33                .Left = LeftScreenLeft + 7        '"Nudge" values found by experimentation
34                .Top = SafeMax(LeftScreenTop, RightScreenTop) + 6
                  'Have seen cases (e.g. the "Bloomberg PC" at Solum) where the line below silently fails to work even _
                   though one can manually set the window width to "stretch across both screens" and the macro recorder _
                   would lead you to believe that the line below should work. :-(
                   
                  'WidthShouldBe = LeftScreenWidth + RightScreenWidth - 25
35                WidthShouldBe = RightScreenLeft + RightScreenWidth - LeftScreenLeft
                   
36                .Width = WidthShouldBe
37                .Height = SafeMin(LeftScreenHeight, RightScreenHeight) - 13
38                If Not ReallyHaveTwoScreens Then
39                    .WindowState = xlMaximized
40                ElseIf .Width < WidthShouldBe - 3 Then
41                    GrowScreen        'see comments in method
42                End If
43            ElseIf NewState = 1 Then
                  'Move to left window
44                .WindowState = xlNormal
45                .Width = LeftScreenWidth
46                .Height = LeftScreenHeight
47                .Left = LeftScreenLeft
48                .Top = LeftScreenTop
49                If ReallyHaveTwoScreens Then .WindowState = xlMaximized
50            End If
51        End With
52        Application.OnRepeat "Repeat Toggle Windows", "ToggleWindow"
53        Exit Sub
ErrHandler:
54        SomethingWentWrong "#ToggleWindow (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
55        Application.EnableEvents = True
56        Application.OnRepeat "Repeat Toggle Window", "ToggleWindow"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeWindowsForVideoRecording
' Author    : Philip Swannell
' Date      : 18-Jan-2017
' Purpose   : Sets windows to size I want for recording videos on PC at home
' -----------------------------------------------------------------------------------------------------------------------
Sub ResizeWindowsForVideoRecording()
          Dim H As Double
          Dim L As Double
          Dim SUH As clsScreenUpdateHandler
          Dim t As Double
          Dim w As Double
          Dim wn As Window
1         On Error GoTo ErrHandler
2         t = 0
3         L = 0
          '    W = 2304 / fX() ' For SCRiPT videos
          '    H = 1296 / fY() 'ratio 16:9
4         w = 1920 / fx()        ' For Cayley videos
5         H = 1080 / fY()        'ratio 16:9

6         Set SUH = CreateScreenUpdateHandler

7         For Each wn In Application.Windows
8             If wn.Visible = True Then
9                 If wn.WindowState <> xlNormal Then
10                    wn.WindowState = xlNormal
11                End If

12                If wn.Top <> 0 Then wn.Top = t
13                If wn.Left <> L Then wn.Left = L
14                If wn.Width <> w Then wn.Width = w
15                If wn.Height <> H Then wn.Height = H

16            End If
17        Next

18        If Application.WindowState <> xlNormal Then Application.WindowState = xlNormal
19        If Application.Top <> t Then Application.Top = t
20        If Application.Left <> L Then Application.Left = L
21        If Application.Width <> w Then Application.Width = w
22        If Application.Height <> H Then Application.Height = H

23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#ResizeWindowsForVideoRecording (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetApplicationCaptions
' Author     : Philip Swannell
' Date       : 08-Jan-2019
' Purpose    : When more than one instance of Excel is open its very hard to tell which workbook is in which instance.
'              This method annotates the application caption so that it's easy to tell.
' -----------------------------------------------------------------------------------------------------------------------
Sub SetApplicationCaptions()
          Dim AllXLs As Collection
          Dim N As Long
          Dim ShowPID As Boolean
          Dim xl As Object
1         On Error GoTo ErrHandler

          Const chPID1 = "Never"
          Const chPID2 = "When multiple Excels open"
          Const chPID3 = "Always"
          Dim PID As Long

2         GetExcelInstances AllXLs
3         N = AllXLs.Count
4         If N = 1 Then
5             Select Case GetSetting(gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID1)
                  Case chPID3
6                     GetWindowThreadProcessId Application.hWnd, PID
7                     ShowPID = True
8                 Case Else
9                     ShowPID = False
10            End Select
              
11            If ShowPID Then
12                Application.caption = "Excel (" + CStr(PID) + ")"
13            Else
14                Application.caption = vbNullString
15            End If
16        Else
17            Select Case GetSetting(gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID1)
                  Case chPID3, chPID2
18                    ShowPID = True
19                Case Else
20                    ShowPID = False
21            End Select
          
22            For Each xl In AllXLs
23                If ShowPID Then
24                    GetWindowThreadProcessId xl.hWnd, PID
25                    xl.caption = "Excel (" + CStr(PID) + ")"
26                Else
27                    xl.caption = vbNullString
28                End If
29            Next xl
              
30        End If
31        Exit Sub
ErrHandler:
32        Throw "#SetApplicationCaptions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetExcelInstances
' Author     : Philip Swannell
' Date       : 27-Jun-2018
' Purpose    : Returns a collection being all of the open Excel instances
'              from https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
'              But changed because code posted there seemed to add each Excel instance to the collection many times - see use of variable AlreadyThere
' -----------------------------------------------------------------------------------------------------------------------
Sub GetExcelInstances(ByRef AllXLs As Collection)
          Dim acc As Object
          Dim guid&(0 To 3)
          Dim hWnd
          Dim hWnd2
          Dim hwnd3
1         On Error GoTo ErrHandler
2         guid(0) = &H20400
3         guid(1) = &H0
4         guid(2) = &HC0
5         guid(3) = &H46000000
          Dim AlreadyThere As Boolean
          Dim xl As Application
6         Set AllXLs = New Collection
7         AllXLs.Add Application 'Ensure "this" application is the first member of the collection
8         Do
9             hWnd = FindWindowExA(0, hWnd, "XLMAIN", vbNullString)
10            If hWnd = 0 Then Exit Do
11            hWnd2 = FindWindowExA(hWnd, 0, "XLDESK", vbNullString)
12            hwnd3 = FindWindowExA(hWnd2, 0, "EXCEL7", vbNullString)
13            If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, guid(0), acc) = 0 Then
14                AlreadyThere = False
15                For Each xl In AllXLs
16                    If TypeName(acc) <> "ProtectedViewWindow" Then
17                        If xl Is acc.Application Then
18                            AlreadyThere = True
19                            Exit For
20                        End If
21                    End If
22                Next
23                If Not AlreadyThere Then
24                    If TypeName(acc) <> "ProtectedViewWindow" Then
25                        AllXLs.Add acc.Application
26                    End If
27                End If
28            End If
29        Loop

30        Exit Sub
ErrHandler:
31        Throw "#GetExcelInstances (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SwitchSheet
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Quickly switch to a worksheet of the active workbook, assigned to Ctrl+Shift+T
' -----------------------------------------------------------------------------------------------------------------------
Sub SwitchSheet()
          Dim Res
          Dim SheetNames
          Dim STK As clsStacker
          Dim ws As Object
          Const Title = "Switch Sheet (" + gAddinName + ")"

1         On Error GoTo ErrHandler
2         If Not ActiveWorkbook Is Nothing Then
3             Set STK = CreateStacker()
4             For Each ws In ActiveWorkbook.Sheets
5                 STK.Stack0D ws.Name
6             Next ws
7             SheetNames = STK.Report
8             SheetNames = sSortedArray(SheetNames)
9             Res = ShowSingleChoiceDialog(SheetNames, , , , , Title, "Select sheet to activate")
10            If Not IsEmpty(Res) Then
11                ActivateSheet CStr(Res)
12            End If
13        End If
14        Application.OnRepeat "Repeat Switch Sheet", "SwitchSheet"
15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#SwitchSheet (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SwitchBook
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Quickly switch to a workbook, including books in other excel instances. Assigned to Ctrl+Shift+B
' -----------------------------------------------------------------------------------------------------------------------
Sub SwitchBook()

          Dim a As Application
          Dim AddinNames
          Dim AllXLs As Collection
          Dim AppNum As Long
          Dim BookName As String
          Dim BookNames
          Dim Categories
          Dim FoundSome As Boolean
          Dim i As Long
          Dim iDRTV As Boolean
          Dim Res
          Dim STK As clsStacker
          Dim STK2 As clsStacker
          Dim TheseBooks
          Dim TopText As String
          
          Const Title = "Switch Book (" + gAddinName + ")"
          
1         On Error GoTo ErrHandler
          
2         Set STK = CreateStacker()
3         iDRTV = InDeveloperMode()
4         GetExcelInstances AllXLs
5         If AllXLs.Count = 1 Then
6             TopText = "Select a workbook to activate"
7             If iDRTV Then
8                 BookNames = WorkbookAndAddInList(1)
                  
9                 If IsEmpty(BookNames) Then
10                    BookNames = CreateMissing()
11                    Categories = CreateMissing()
12                Else
13                    Categories = sReshape("Workbook", sNRows(BookNames), 1)
14                End If
15                AddinNames = WorkbookAndAddInList(2)
16                If IsEmpty(AddinNames) Then
17                    AddinNames = CreateMissing()
18                Else
19                    Categories = sArrayStack(Categories, sReshape("Addin", sNRows(AddinNames), 1))
20                End If
21                BookNames = sArrayStack(BookNames, AddinNames)

                  Dim BooksAndCategories
22                BooksAndCategories = sArrayRange(BookNames, Categories)
23                BooksAndCategories = sSortedArray(BooksAndCategories)
24                BookNames = sSubArray(BooksAndCategories, 1, 1, , 1)
25                Categories = sSubArray(BooksAndCategories, 1, 2, , 1)
26            Else
27                If Application.Workbooks.Count = 0 Then GoTo EarlyExit
28                BookNames = sSortedArray(WorkbookAndAddInList(1))
29                Categories = CreateMissing()
30            End If
31            Res = ShowSingleChoiceDialog(BookNames, , Categories, , , Title, TopText, , "Type:", , , "Workbook")
32            If IsEmpty(Res) Then GoTo EarlyExit
33            ActivateBook CStr(Res)
34        Else ' Multiple Excels
35            Set STK2 = CreateStacker()
36            TopText = "There are " + CStr(AllXLs.Count) + " Excel instances running with" & vbLf & _
                  "instance numbers shown next to each" & vbLf & _
                  "workbook." & vbLf & "Select a workbook to activate."
37            i = 0
38            For Each a In AllXLs
39                i = i + 1
40                TheseBooks = WorkbookAndAddInList(1, a)
41                If Not IsEmpty(TheseBooks) Then
42                    FoundSome = True
43                    STK.Stack2D sArrayConcatenate(CStr(i) & " ", TheseBooks)
44                    STK2.Stack2D sReshape("WorkBook", sNRows(TheseBooks), 1)
45                End If
46            Next a
47            Categories = CreateMissing()
48            If iDRTV Then
49                i = 0
50                For Each a In AllXLs
51                    i = i + 1
52                    TheseBooks = WorkbookAndAddInList(2, a)
53                    If Not IsEmpty(TheseBooks) Then
54                        FoundSome = True
55                        STK.Stack2D sArrayConcatenate(CStr(i) & " ", TheseBooks)
56                        STK2.Stack2D sReshape("AddIn", sNRows(TheseBooks), 1)
57                    End If
58                    Exit For 'because too confusing to give ability to jump to addins in other Excel instances
59                Next a
60            End If

61            If Not FoundSome Then GoTo EarlyExit
62            BookNames = STK.Report
63            If iDRTV Then
64                Categories = STK2.Report
65                BooksAndCategories = sArrayRange(BookNames, Categories)
66                BooksAndCategories = sSortedArray(BooksAndCategories)
67                BookNames = sSubArray(BooksAndCategories, 1, 1, , 1)
68                Categories = sSubArray(BooksAndCategories, 1, 2, , 1)
69            Else
70                BookNames = sSortedArray(BookNames)
71                Categories = CreateMissing()
72            End If

73            Res = ShowSingleChoiceDialog(BookNames, , Categories, , , Title, TopText, , "Type:", , , "Workbook")
74            If IsEmpty(Res) Then GoTo EarlyExit
75            BookName = sStringBetweenStrings(Res, " ")
76            AppNum = CLng(sStringBetweenStrings(Res, , " "))
77            ActivateBook AllXLs(AppNum).Workbooks(BookName)
78        End If
EarlyExit:
79        If AppNum <= 1 Then
80            Application.OnRepeat "Repeat Switch Book", "SwitchBook"
81        Else
82            Application.OnRepeat vbNullString, vbNullString
83        End If
84        Exit Sub
ErrHandler:
85        SomethingWentWrong "#SwitchBook (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub

