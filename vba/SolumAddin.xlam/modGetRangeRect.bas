Attribute VB_Name = "modGetRangeRect"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modMain
' Author    : Philip Swannell - downloaded from
'             http://dailydoseofexcel.com/archives/2007/08/30/positioning-a-userform-over-a-cell/
' Date      : 17-Oct-2013
' Purpose   : Method GetRangeRect will allow the positioning of forms relative to a cell
'             7 April 2016. Commented out code that's not required in Excel versions >= 10 i.e. >= Excel 2002
'             19-June-2017. Downloaded a new version (see link "Download RangePos Beta3.zip") from web page refered to above
'             it seems that support for multi-monitor setups is improved
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type POINTAPI
    x As Long
    y As Long
End Type

'PGS 24/11/2015 used Microsoft Office 2010 Code Compatibility Inspector to make changes for 64-bit compatibility

Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "USER32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "USER32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "USER32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "USER32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
'    'TODO PGS 24/11/15 Microsoft Code Combatibility Inspector did not suggest a correction to the 32-bit declaration below, so the change I made is guess-work :-(
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As LongPtr) As Long
Public Declare PtrSafe Function GetForegroundWindow Lib "USER32" () As LongPtr
' -----------------------------------------------------------------------------------------------------------------------
'Main Functions
' -----------------------------------------------------------------------------------------------------------------------
Public Function GetRangeRect(rSelection As Excel.Range) As RECT
          Dim rVisible As Excel.Range
          Dim wnd As Excel.Window

          Dim iPane As Long
          Dim PT As POINTAPI
          Dim RC As RECT

1         On Error GoTo errH:

2         Set wnd = rSelection.Worksheet.Parent.Windows(1)
3         If PanesAreSwapped(wnd) Then PanesReorder wnd

4         iPane = PaneSelection(wnd, rSelection, rVisible)
5         If iPane = 0 Then Err.Raise vbObjectError + &H1000, "GetRangeRect", "GetRangeRect: Range not visible"

6         PT = PaneOrigin(wnd)
7         With wnd
8             If .FreezePanes Then
                  'we have to work from the middle iso the topleft of the activepane.
9                 If .SplitColumn > 1 And (iPane = 1 Or iPane = 3) Then
10                    PT.x = PT.x - fx * .Panes(1).VisibleRange.Width * .Zoom / 100
11                End If
12                If .SplitRow > 1 And (iPane = 1 Or (iPane = 2 And .Panes.Count > 2)) Then
13                    PT.y = PT.y - fY * .Panes(1).VisibleRange.Height * .Zoom / 100
14                End If
15            End If
16        End With

17        RC.Left = PT.x: If rVisible.Column < rSelection.Column Then RC.Left = RC.Left + RangePixelsWidth(rVisible.Resize(, rSelection.Column - rVisible.Column))
18        RC.Top = PT.y: If rVisible.row < rSelection.row Then RC.Top = RC.Top + RangePixelsHeight(rVisible.Resize(rSelection.row - rVisible.row))

          'this may partially extend over a split or the window edge
19        RC.Right = RC.Left + RangePixelsWidth(rSelection)
20        RC.Bottom = RC.Top + RangePixelsHeight(rSelection)

21        GetRangeRect = RC
errH:
End Function

Private Function PaneOrigin(wnd As Window) As POINTAPI
          ' Returns the position of the upperleft corner of the active pane

          ' Complexities:

          ' Get.Cell returns wrong headersizes if zoom is not 100. Where possible we use the move of SplitVert/SplitHorz when Displayheadings are toggled.
          ' Get.Cell returns wrong headerwidth if activepane has lower row magnitude than pane on other side of vertical split.

          ' SplitVert returns 0 if Pane1 scrollrow + visiblerange.rows.count = rows.count+1
          ' SplitHorz returns 0 if Pane1 scrollcol + visiblerange.cols.count = cols.count+1

          ' Known issues:
          ' With 2 pane splits small inaccuracies can occur when zoom <> 100
          ' but we've taken care of most exceptions. :)

          Dim fmlX As String
          Dim fmlY As String
          Dim RC(0 To 1) As RECT

          Dim dh(0 To 1) As Double
          Dim dv(0 To 1) As Double
          Dim dx(0 To 1) As Double
          Dim dy(0 To 1) As Double

          Dim bHead As Boolean  'true if DisplayHeadings is on
          Dim bOutl As Boolean  'true if DisplayOutline  is on

          Dim bRows As Boolean  'true if pane has row headings
          Dim bCols As Boolean  'true if pane has col headings
          Dim bSwap As Boolean  'true if pane has row magnitude problem
          Dim bDirt As Boolean  'true for temp splits

1         Application.ScreenUpdating = False

2         With wnd
3             If PanesAreSwapped(wnd) Then PanesReorder wnd

4             bHead = .DisplayHeadings
5             bOutl = .DisplayOutline

6             fmlX = "GET.CELL(42," & .ActivePane.VisibleRange.address(1, 1, xlR1C1, 1) & ")"
7             fmlY = "GET.CELL(43," & .ActivePane.VisibleRange.address(1, 1, xlR1C1, 1) & ")"

8             If .FreezePanes Then
                  'SplitH/SplitV do not move on changing DisplayHeadings.
                  'For 2 pane splits we have to rely on less precise zoom calculation.
9                 .DisplayHeadings = False
10                .DisplayOutline = False
11                dx(0) = ExecuteExcel4Macro2(fmlX)
12                dy(0) = ExecuteExcel4Macro2(fmlY)

                  'Size of outline
13                .DisplayOutline = bOutl
14                dx(1) = ExecuteExcel4Macro2(fmlX) - dx(0)
15                dy(1) = ExecuteExcel4Macro2(fmlY) - dy(0)
16                .DisplayOutline = False

                  'Size of headers
17                .DisplayHeadings = bHead
18                dh(0) = ExecuteExcel4Macro2(fmlX) - dx(0)
19                dv(0) = ExecuteExcel4Macro2(fmlY) - dy(0)

20                .DisplayOutline = bOutl
                  'Adjust header for zoom error
21                If .SplitHorizontal Then dx(1) = dx(1) + dh(0) Else dx(1) = dx(1) + dh(0) * .Zoom / 100
22                If .SplitVertical Then dy(1) = dy(1) + dv(0) Else dy(1) = dy(1) + dv(0) * .Zoom / 100

23            Else

                  'Get the base values (excluding DisplayHeadings but including optional OutlineHeadings)
24                .DisplayHeadings = False
25                If Not .Split And bHead And .Zoom <> 100 Then
                      'No splits: create them
26                    If .ScrollColumn + .VisibleRange.Columns.Count <= .ActiveSheet.Columns.Count And _
                          .ScrollRow + .VisibleRange.Rows.Count <= .ActiveSheet.Rows.Count Then
27                        bDirt = True
28                        .SplitHorizontal = .UsableWidth + 1
29                        .SplitVertical = .UsableHeight + 1
30                    End If
31                End If

32                dx(0) = ExecuteExcel4Macro2(fmlX)
33                dy(0) = ExecuteExcel4Macro2(fmlY)
34                dh(0) = .SplitHorizontal
35                dv(0) = .SplitVertical

36                If bHead Then
37                    .DisplayHeadings = True
38                    dh(1) = .SplitHorizontal
39                    dv(1) = .SplitVertical

40                    bRows = (dx(0) >= 0 And dx(0) < dh(0)) Or ((dh(0) = 0 Or dh(1) = 0) And dx(0) >= 0 And dx(0) < .Panes(1).VisibleRange.Width)
41                    bCols = (dy(0) >= 0 And dy(0) < dv(0)) Or ((dv(0) = 0 Or dv(1) = 0) And dy(0) >= 0 And dy(0) < .Panes(1).VisibleRange.Height)

42                    If bRows And .Split And Not .FreezePanes Then
                          'Swap if 'other' pane has 'wider' row headings.
43                        bSwap = Len(Format$(.ActivePane.VisibleRange.Rows(.ActivePane.VisibleRange.Rows.Count).row, "000")) < _
                              Len(Format$(.Panes(1 + ((.ActivePane.index + .Panes.Count \ 2) - 1) Mod .Panes.Count).VisibleRange.Rows(.Panes(1 + ((.ActivePane.index + .Panes.Count \ 2) - 1) Mod .Panes.Count).VisibleRange.Rows.Count).row, "000"))
44                    End If
45                    If bRows And bSwap Then
                          'recompute dx(1) for the pane on the other side of the vertical split aka the horizontal bar.
46                        .Panes(1 + ((.ActivePane.index + .Panes.Count \ 2) - 1) Mod .Panes.Count).Activate
47                        dx(1) = ExecuteExcel4Macro2(fmlX) - dx(0)
48                        .Panes(1 + ((.ActivePane.index + .Panes.Count \ 2) - 1) Mod .Panes.Count).Activate:
49                        .ActivePane.ScrollRow = .ActivePane.ScrollRow
50                    ElseIf bRows Then
51                        dx(1) = ExecuteExcel4Macro2(fmlX) - dx(0)
52                    End If
53                    If bCols Then dy(1) = ExecuteExcel4Macro2(fmlY) - dy(0)

54                    If bRows And dh(1) > 0 And dh(1) < dh(0) Then
55                        dx(1) = dh(0) - dh(1)
56                    ElseIf bRows Then
                          'inexact when zoomed.
57                        dx(1) = dx(1) * .Zoom / 100
58                    End If
59                    If bCols And dv(1) > 0 And dv(1) < dv(0) Then
60                        dy(1) = dv(0) - dv(1)
61                    ElseIf bCols Then
                          'inexact when zoomed!
62                        dy(1) = dy(1) * .Zoom / 100
63                    End If
64                    If dx(1) < 0 Then dx(1) = 0
65                    If dy(1) < 0 Then dy(1) = 0

66                End If

67                If bDirt Then .Split = False

68            End If
69        End With

70        Application.ScreenUpdating = True

          'Rectangle coordinates
          Dim WindowHandle As Variant    'Long or LongPtr
71        WindowHandle = xlWindowHandle(wnd)
          'PGS 21 Jun 2017. This is an almighty bodge. I don't understand why xlWindowHandle is sometimes failing (returning a handle of zero). In practice we're almost always _
           interested in the ActiveWindow, which is also (assuming Excel is the active application) the ForegroundWindow
72        If WindowHandle = 0 Then WindowHandle = GetForegroundWindow()

73        GetWindowRect WindowHandle, RC(0)
74        GetClientRect WindowHandle, RC(1)

75        PaneOrigin.x = fx * (dx(0) + dx(1)) + RC(0).Left + (RC(0).Right - RC(0).Left - RC(1).Right) \ 2
76        PaneOrigin.y = fY * (dy(0) + dy(1)) + RC(0).Bottom - RC(1).Bottom - (RC(0).Right - RC(0).Left - RC(1).Right) \ 2
End Function

Private Function PaneSelection(wnd As Excel.Window, ByRef rSel As Range, ByRef rVis As Range) As Long
          'finds the pane where rSel is best visible
          'returns: ActivePane index if selection completely visible in multiple panes.
          '         0 if not visible in window

          'Note:
          'sets rSel to intersect of range selection and selected pane's visible range
          'sets rVis to selected pane's visible range

          Dim aRng(1 To 4) As Range
          Dim M&
          Dim mCnt&
          Dim N&

1         With wnd
2             For N = 1 To .Panes.Count
3                 Set aRng(N) = Intersect(rSel, .Panes(N).VisibleRange)
4                 If Not aRng(N) Is Nothing Then
5                     If aRng(N).Count > mCnt Then
6                         M = N: mCnt = aRng(N).Count
7                     ElseIf aRng(N).Count = mCnt And N = .ActivePane.index Then
8                         M = N
9                     End If
10                End If
11            Next
12            If M Then
13                Set rSel = aRng(M)
14                Set rVis = .Panes(M).VisibleRange
15                PaneSelection = M
16            End If
17        End With
End Function

Private Function PanesAreSwapped(wnd As Excel.Window) As Boolean
          'Returns true if pane2 is NorthEast in a 4pane window)
          Dim aRng(1 To 4) As Excel.Range
          Dim dAdj#
          Dim N&

1         With wnd
2             If .Panes.Count = 4 Then
3                 For N = 1 To 4: Set aRng(N) = .Panes(N).VisibleRange: Next
4                 If .FreezePanes Then
5                     PanesAreSwapped = aRng(1).Column <> aRng(3).Column
6                 Else

7                     If aRng(1).row = aRng(4).row And aRng(1).Column = aRng(4).Column Then
8                         If aRng(1).Height = aRng(4).Height And aRng(1).Width = aRng(4).Width Then
                              'totally square. we must move a split to see which is which
9                             If .SplitHorizontal >= 50 Then dAdj = -40 Else dAdj = 40
10                        End If
11                    End If

12                    If dAdj Then .SplitHorizontal = .SplitHorizontal + dAdj
13                    PanesAreSwapped = (aRng(3).Height = aRng(1).Height And aRng(3).Width = aRng(4).Width)
14                    If dAdj Then .SplitHorizontal = .SplitHorizontal - dAdj

15                End If
16            End If
17        End With
End Function

Private Sub PanesReorder(wnd As Excel.Window)
          'Forces the panes to the default sequence

          'Excel always expect the panes in NW/NE/SW/SE order.
          'However when you manually drag the splitbars it is possible to create a panes
          'collection where the VisibleRange of panes 2 and 3 are reversed.

          Dim bFP As Boolean
          Dim dSV As Double
          Dim iPane As Long
          Dim lSC(0 To 1) As Long
          Dim lSR(0 To 1) As Long
          Dim rCell As Range
          Dim rSele As Range

1         With wnd
2             If .Panes.Count = 4 Then
                  'Store info
3                 Set rCell = .ActiveCell
4                 Set rSele = .RangeSelection
5                 iPane = .ActivePane.index
6                 bFP = .FreezePanes
7                 lSR(0) = .Panes(1).ScrollRow
8                 lSR(1) = .Panes(4).ScrollRow
9                 lSC(0) = .Panes(1).ScrollColumn
10                lSC(1) = .Panes(4).ScrollColumn

11                Do While .SplitVertical < 1
                      'avoid bug when rows are scrolled 'beyond'
12                    .Panes(1).ScrollRow = .Panes(1).ScrollRow - 1
13                Loop
                  'Ensure Vertical is set after Horizontal
14                dSV = .SplitVertical
15                .SplitVertical = 0
16                .SplitVertical = dSV

                  'Restore info
17                If bFP Then .FreezePanes = True Else .Panes(iPane).Activate
18                rSele.Select
19                rCell.Activate
20                .Panes(1).ScrollRow = lSR(0)
21                .Panes(4).ScrollRow = lSR(1)
22                .Panes(1).ScrollColumn = lSC(0)
23                .Panes(4).ScrollColumn = lSC(1)

24            End If
25        End With
End Sub

'Compute width/height per cell to avoid rounding errors.
Private Function RangePixelsWidth(rRange As Range) As Long
          Dim rCell As Range
1         For Each rCell In rRange.Columns
2             RangePixelsWidth = RangePixelsWidth + Application.WorksheetFunction.Round(fx * ActiveWindow.Zoom / 100 * rCell.Width, 0)
3         Next
End Function
Private Function RangePixelsHeight(rRange As Range) As Long
          Dim rCell As Range
1         For Each rCell In rRange.Rows
2             RangePixelsHeight = RangePixelsHeight + Application.WorksheetFunction.Round(fY * ActiveWindow.Zoom / 100 * rCell.Height, 0)
3         Next
End Function

'ScreenResolution
Function fx() As Double
1         fx = ScreenDPI(0) / 72
End Function

Function fY() As Double
1         fY = ScreenDPI(1) / 72
End Function

Function ScreenDPI(bVert As Boolean) As Long
          Static lDpi(1) As Long
          Static hdc As Variant    'Long or LongPtr
1         If lDpi(0) = 0 Then
2             hdc = GetDC(0)
3             lDpi(0) = GetDeviceCaps(hdc, 88&)                           'horz
4             lDpi(1) = GetDeviceCaps(hdc, 90&)                           'vert
5             hdc = ReleaseDC(0, hdc)
6         End If
7         ScreenDPI = lDpi(Abs(bVert))
End Function

'Handles
Private Function xlApplicationHandle() As Long
          Static H As Long
1         If H = 0 Then
2             H = Application.hWnd
3         End If
4         xlApplicationHandle = H
End Function

Private Function xlDesktopHandle() As Long
          Static H As Variant    'Long or LongPtr
1         If H = 0 Then H = FindWindowEx(xlApplicationHandle, 0&, "XLDESK", vbNullString)
2         xlDesktopHandle = H
End Function

Private Function xlWindowHandle(wnd As Window) As Long
          Dim H As Variant    'Long or LongPtr
1         H = FindWindowEx(xlDesktopHandle, 0, "EXCEL7", wnd.caption)
2         If H = 0 Then H = WindowSearch(xlDesktopHandle, "EXCEL7", wnd.caption & "*")
3         xlWindowHandle = H
End Function

Private Function WindowSearch(ByVal hTop As Long, ByVal sClass As String, ByVal sCaptionPattern As String) As Long
          Dim hWnd As Variant    'Long or LongPtr
          Dim lLen As Long
          Dim sBuf As String
1         sBuf = String(&HFF&, 0)
2         Do
3             hWnd = FindWindowEx(hTop, hWnd, sClass, vbNullString)
4             lLen = GetWindowText(hWnd, StrPtr(sBuf), &HFF&)
5         Loop Until hWnd = 0 Or LCase$(Left$(sBuf, lLen)) Like LCase$(sCaptionPattern)
6         WindowSearch = hWnd
End Function
