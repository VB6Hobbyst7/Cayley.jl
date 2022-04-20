Attribute VB_Name = "modGroupingButtons"
Option Explicit

Private Const FontForMore = "Wingdings 3"
Private Const MoreCharNo = 125
Private Const MoreCharSimple = ">"
Private Const FontForLess = "Wingdings 3"
Private Const LessCharNo = 124
Private Const LessCharSimple = "<"

Private Function GetButtonStatus(b As Button) As Boolean
          Dim ShowMore As String
1         On Error GoTo ErrHandler
2         If sFontIsInstalled(FontForMore) Then
3             ShowMore = " " + Chr$(MoreCharNo)
4         Else
5             ShowMore = " " + MoreCharSimple
6         End If
7         GetButtonStatus = b.caption = ShowMore
8         Exit Function
ErrHandler:
9         Throw "#GetButtonStatus (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub SetButtonStatus(b As Button, Expand As Boolean)
          Dim ShowLess As String
          Dim ShowMore As String

1         On Error GoTo ErrHandler
2         If Expand Then
3             If sFontIsInstalled(FontForMore) Then
4                 ShowMore = " " + Chr$(MoreCharNo)
5                 b.Font.Name = FontForMore
6             Else
7                 ShowMore = " " + MoreCharSimple
8                 b.Font.Name = "Calibri"
9             End If
10            b.caption = ShowMore
11        Else
12            If sFontIsInstalled(FontForLess) Then
13                ShowLess = " " + Chr$(LessCharNo)
14                b.Font.Name = FontForLess
15            Else
16                ShowLess = " " + LessCharSimple
17                b.Font.Name = "Calibri"
18            End If
19            b.caption = ShowLess
20        End If

21        b.HorizontalAlignment = xlHAlignLeft
22        b.Font.ColorIndex = 48
23        b.Placement = xlMoveAndSize

24        Exit Sub
ErrHandler:
25        Throw "#SetButtonStatus (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GroupingButton
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Buttons for hiding and showing columns - alternative to Excel's Outlining feature
' -----------------------------------------------------------------------------------------------------------------------
Sub GroupingButton()
          Dim b As Button
1         On Error GoTo ErrHandler
2         If VarType(Application.Caller) = vbString Then
3             Set b = ActiveSheet.Buttons(Application.Caller)
4         ElseIf sElapsedTime() - LastAltBacktickTime < 0.5 Then
5             Set b = LastAltBacktickButton
6         End If
7         If Not IsShiftKeyDown Then
8             GroupingButtonCore b
9             If MultipleButtonsExist(b.Parent) Then
10                TemporaryMessage "Tip: Shift-click does all column groups at once!", , False
11            End If
12        Else        'If the Shift Key is down then expand\collapse all grouping buttons on the sheet
13            GroupingButtonDoAllOnSheet ActiveSheet, GetButtonStatus(b)
14        End If

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#GroupingButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GroupingButtonCore
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : "Core" code - can be called either from GroupingButton that should be the
'              OnAction of each button or else from GroupingButtonDoAllOnSheet
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GroupingButtonCore(b As Button, Optional Expand As Variant)
1         On Error GoTo ErrHandler
          Dim R As Range
          Dim SPH As Object

2         Set SPH = CreateSheetProtectionHandler(b.Parent)

3         Set R = RangeBeneathGroupingButton(b)

4         If VarType(Expand) <> vbBoolean Then Expand = GetButtonStatus(b)

5         With R
6             If Expand Then
7                 .EntireColumn.Hidden = False
8                 SetButtonStatus b, False
9             Else
10                If .Columns.Count > 1 Then
11                    .Offset(, 1).Resize(, .Columns.Count - 1).EntireColumn.Hidden = True
12                    SetButtonStatus b, True
13                End If
14            End If
15        End With

16        b.Parent.Activate
17        Exit Sub
ErrHandler:
18        Throw "#GroupingButtonCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeBeneathGroupingButton
' Author    : Philip Swannell
' Date      : 7 Nov 2016
' Purpose   : It's difficult to determine the correct range associated with a grouping button
'             in a way that works on all versions of Excel.
'             In Excel 2013 and 2016 Range(b.TopLeftCell, b.BottomRightCell.Offset(, -1)) would suffice
'             but in Excel 2010 TopLeftCell and BottomRightCell properties don't work so well, particularly
'             in relation to hidden columns.
' -----------------------------------------------------------------------------------------------------------------------
Private Function RangeBeneathGroupingButton(b As Button) As Range
          Dim bb As Double
          Dim bl As Double
          Dim br As Double
          Dim bt As Double
          Dim c As Range
          Dim Cell1 As Range
          Dim Cell2 As Range
          Dim Dist As Double
          Dim i As Long
          Dim MinDist As Double
          Dim RangeToSearch As Range
          Dim RealBottomRight As Range
          Dim RealTopLeft As Range
          Dim StartCell As Range
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         If b.Width = 0 Then
              'Suspect the line below won't work on Office 2010, but on the other hand should not hit this situation in practice,
              'not least because it's not so easy to click a zero-width button...
3             Set RangeBeneathGroupingButton = b.TopLeftCell.Parent.Range(b.TopLeftCell, b.BottomRightCell.Offset(, -1))
4             Exit Function
5         End If

6         bt = b.Top
7         bl = b.Left
8         br = bl + b.Width
9         bb = bt + b.Height

10        Set StartCell = b.TopLeftCell
11        Set ws = StartCell.Parent
12        Set Cell1 = StartCell.Offset(IIf(StartCell.row = 1, 0, -1), IIf(StartCell.Column = 1, 0, -1))
13        Set Cell2 = StartCell.Offset(IIf(StartCell.row = ws.Rows.Count, 0, 1), IIf(StartCell.Column = ws.Columns.Count, 0, 1))

14        Set RangeToSearch = ws.Range(Cell1, Cell2)
15        MinDist = 100000000
16        For Each c In RangeToSearch.Cells
17            Dist = (c.Top - bt) ^ 2 + (c.Left - bl) ^ 2
18            If Dist < MinDist Then
19                Set RealTopLeft = c
20                MinDist = Dist
21            End If
22        Next

          'Assume that first column cannot have zero width....
23        Do While RealTopLeft.Width = 0 And RealTopLeft.Column < ws.Columns.CountLarge
24            Set RealTopLeft = RealTopLeft.Offset(, 1)
25        Loop

26        For i = 0 To ws.Columns.CountLarge - RealTopLeft.Column
27            If RealTopLeft.Offset(0, i).Left > br + 0.00001 Then        'Have seen cases when there should have been equality but the LHS exceeded the RHS by 1e-14
28                Exit For
29            End If
30        Next i
31        If i <= 2 Then
32            Set RealBottomRight = RealTopLeft
33        Else
34            Set RealBottomRight = RealTopLeft.Offset(0, i - 2)
35        End If

36        Set RangeBeneathGroupingButton = ws.Range(RealTopLeft, RealBottomRight)

37        Exit Function
ErrHandler:
38        Throw "#RangeBeneathGroupingButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GroupingButtonDoAllOnSheet
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Expand or Collapse all of the groups of columns on a sheet that have grouping buttons on them
' -----------------------------------------------------------------------------------------------------------------------
Sub GroupingButtonDoAllOnSheet(ws As Worksheet, Expand As Boolean)
          Dim b As Button
1         Application.ScreenUpdating = False

2         On Error GoTo ErrHandler
3         For Each b In ws.Buttons
4             If InStr(b.OnAction, "GroupingButton") > 0 Then
5                 GroupingButtonCore b, Expand
6             End If
7         Next
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#GroupingButtonDoAllOnSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MultipleButtonsExist
' Author    : Philip Swannell
' Date      : 05-Oct-2015
' Purpose   : Returns True if more than one Grouping button is on the sheet
' -----------------------------------------------------------------------------------------------------------------------
Private Function MultipleButtonsExist(ws As Worksheet) As Boolean
          Dim b As Button
          Dim N As Long

1         For Each b In ws.Buttons
2             If InStr(b.OnAction, "GroupingButton") > 0 Then
3                 N = N + 1
4                 If N > 1 Then
5                     MultipleButtonsExist = True
6                     Exit Function
7                 End If
8             End If
9         Next
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddGroupingButtons
' Author    : Philip Swannell
' Date      : 30-Sep-2015
' Purpose   : Adds grouping button to the top row of each area of the current selection
' -----------------------------------------------------------------------------------------------------------------------
Sub AddGroupingButtons()
          Dim Area As Range
1         If TypeName(Selection) <> "Range" Then Exit Sub
2         On Error GoTo ErrHandler
3         If Not UnprotectAsk(Selection.Parent, "Add Grouping Buttons") Then Exit Sub
4         For Each Area In Selection.Areas
5             AddGroupingButtonToRange Area, True
6         Next Area
7         Application.OnRepeat "Repeat Add Grouping Button", "AddGroupingButtons"
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#AddGroupingButtons (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddGroupingButtonToRange
' Author    : Philip Swannell
' Date      : 05-Oct-2016
' Purpose   : What it says on the tin, can be called from external methods
' -----------------------------------------------------------------------------------------------------------------------
Sub AddGroupingButtonToRange(R As Range, Expand As Boolean)
          Dim b As Button
          Dim c As Range
1         On Error GoTo ErrHandler
2         Set c = R.Rows(1)
3         Set b = c.Parent.Buttons.Add(c.Left, c.Top, c.Width, c.Height)
4         SetButtonStatus b, False
5         b.OnAction = "GroupingButton"
6         If Not Expand Then
7             GroupingButtonCore b, False
8         End If

9         Exit Sub
ErrHandler:
10        Throw "#AddGroupingButtonToRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
