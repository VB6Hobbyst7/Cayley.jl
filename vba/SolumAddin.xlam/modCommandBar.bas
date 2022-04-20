Attribute VB_Name = "modCommandBar"
Option Explicit
Private m_UsersChoice As Long
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowCommandBarPopup
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Constructs a (temporary) command-bar popup menu, displays it, and returns the item
'             that the user selected either as an Index (base 1) or as the string selected.
' Arguments:
' TheChoices. In the simplest case a column array of strings in which case a menu
'             with no sub-menus appears. To see sub menus a multi column array has to be
'             passed for example:
'
'         Animal       Ca&t
'                      Dog
'         --Vegetable  Carrot
'                      Turnip
'         Mineral
'
'             In this example, the top level menu has three elements: Animal, Vegetable, Mineral
'             and the first two elements each have two-element sub-menus Cat, Dog and
'             Carrot, Turnip respectively. There is a BeginGroup horizontal line before
'             Vegetable, because of the double minus sign and the the "t" of Cat is the
'             accelerator key for that menu element thanks to the ampersand character.
' FaceIDs     A column array with the same number of rows as TheChoices. These determine the icons
'             that appear to the left of each menu element. Allowed formats:
'          1) a FaceID number - see workbook AllFaceIDs
'          2) a ImageMso as recognised by Application.CommandBars.GetImageIso - see ??? for a complete list
'          3) the address of file containing a 16 x 16 bitmap, must end in .bmp and ideally there should be
'             a companion "Mask" file with the name <FileName>Mask.bmp
'          4) The name of a 16x16 pixel picture held on the sheet "Custom Icons". Ideally there should be
'             a companion "Mask" picture with the name <PictureName>Mask
' EnableFlags A column array of Booleans to indicate which menu elements are to be enabled. If omitted all
'             elements are enabled.
' CurrentChoice Is passed as either an index (i.e. row number) or as text matching an element of TheChoices
'             CurrentChoice can either be embellished (with & characters) or it can not be embellished.
' AnchorObject    Can be passed as a Range in which case the menu appears with its top left corner at the centre
'             of that range. For added control can be passed as a clsPositionInstructions
'             If AnchorObject is omitted then the menu appears near the mouse pointer unless the user has pressed the
'             right-click key (or Shift F10) within the last 0.5 seconds in which case the menu appears near the active cell
' ReturnIndex Boolean. If True then the return is an index (i.e. row number) of the user's selection or zero
'             if the user escapes out of the menu. If False then the return is a string, giving the (unembellished)
'             caption that the user selected in the menu or the string "#Cancel!" if they escaped out of the menu.
' -----------------------------------------------------------------------------------------------------------------------
Function ShowCommandBarPopup(ByVal TheChoices As Variant, _
        Optional ByVal FaceIDs As Variant, _
        Optional EnableFlags As Variant, _
        Optional CurrentChoice As Variant, _
        Optional ByVal AnchorObject As Object, _
        Optional ReturnIndex As Boolean)

          Dim i As Long
          Dim isFromKeyboard As Boolean
          Dim j As Long
          Dim NR As Long
          Dim Res
          Dim RightMosts As Variant
          Static TempCommandBar As Office.CommandBar
          Dim TickThisOne As Boolean
          Dim x As Double
          Dim y As Double

1         On Error GoTo ErrHandler

2         isFromKeyboard = (sElapsedTime() - SafeMax(LastShiftF10Time, LastAltBacktickTime)) < 0.5

          'Error check the inputs
3         NR = sNRows(TheChoices)
4         If IsMissing(FaceIDs) Then FaceIDs = sReshape(0, NR, 1)
5         If IsMissing(EnableFlags) Then EnableFlags = sReshape(True, NR, 1)
6         RightMosts = sReshape(vbNullString, NR, 1)
7         If sNRows(FaceIDs) <> NR Or sNCols(FaceIDs) <> 1 Then
8             Throw "FaceIDs must have one column and same number of rows as TheChoices"
9         End If
10        If sNRows(EnableFlags) <> NR Or sNCols(EnableFlags) <> 1 Then
11            Throw "EnableFlags must have one column and same number of rows as TheChoices"
12        End If
13        Force2DArray TheChoices
14        Force2DArray EnableFlags
15        Force2DArray FaceIDs
16        Force2DArray RightMosts
17        For i = 1 To NR
18            If VarType(EnableFlags(i, 1)) <> vbBoolean Then
19                Throw "EnableFlags must be a column array with Boolean elements"
20            End If
21        Next i

22        m_UsersChoice = 0

          'Populate a column RightMosts that holds the menu items that the user can select _
           i.e. the CommandBarButtons rather than the CommandBarPopups.
23        For i = 1 To NR
24            For j = sNCols(TheChoices) To 1 Step -1
25                If IsValidChoice(TheChoices(i, j)) Then
26                    RightMosts(i, 1) = TheChoices(i, j)
27                    Exit For
28                End If
29            Next j
30        Next i

          'Deal with CurrentChoice by overwriting any FaceID passed in with a FaceId that will render as a tick
31        If Not IsMissing(CurrentChoice) Then
32            For i = 1 To NR
33                TickThisOne = False
34                If VarType(CurrentChoice) = vbString Then
35                    If LCase$(CurrentChoice) = LCase$(RightMosts(i, 1)) Then
36                        TickThisOne = True
37                    ElseIf LCase$(CurrentChoice) = LCase$(Unembellish(CStr(RightMosts(i, 1)))) Then
38                        TickThisOne = True
39                    End If
40                ElseIf CurrentChoice = i Then
41                    TickThisOne = True
42                End If
43                If TickThisOne Then
44                    FaceIDs(i, 1) = 1087
45                End If
46            Next i
47        End If

          'Method PopulateCommandBar is quite slow (~ 0.75 seconds for 44 element menu) so do not re construct if we don't have to...
          Static PrevChoices, PrevEnableFlags, PrevFaceIds, RepopCommandBar As Boolean
48        RepopCommandBar = True
49        If Not TempCommandBar Is Nothing Then
50            If Not IsEmpty(PrevChoices) Then
51                If sArraysIdentical(TheChoices, PrevChoices) Then
52                    If sArraysIdentical(EnableFlags, PrevEnableFlags) Then
53                        If sArraysIdentical(FaceIDs, PrevFaceIds) Then
54                            RepopCommandBar = False
55                        End If
56                    End If
57                End If
58            End If
59        End If

TryAgain:
60        If RepopCommandBar Then
61            SafeDeleteCommandBar TempCommandBar
62            Set TempCommandBar = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
63            PopulateCommandBar TempCommandBar, TheChoices, EnableFlags, FaceIDs
64            PrevChoices = TheChoices
65            PrevEnableFlags = EnableFlags
66            PrevFaceIds = FaceIDs
67        End If

68        If AnchorObject Is Nothing Then
69            If Not ActiveCell Is Nothing Then
70                If isFromKeyboard Then
71                    Set AnchorObject = ActiveCell
72                End If
73            End If
74        End If
          Dim HaveTriedAgain As Boolean, CopyOfErr As String

          'Figure out where the menu should go.
75        XYCoordinatesOfObjectCentre AnchorObject, x, y

76        On Error GoTo PopupError 'Rather ugly error handling in this method...
77        If x <> 0 And y <> 0 Then
78            TempCommandBar.ShowPopup x, y
79        Else
80            TempCommandBar.ShowPopup
81        End If
82        If False Then
PopupError:
83            CopyOfErr = Err.Description
84            If HaveTriedAgain Then Throw CopyOfErr
85            HaveTriedAgain = True
86            RepopCommandBar = True
87            GoTo TryAgain
88        End If

89        Application.ScreenUpdating = True

90        If m_UsersChoice <> 0 Then
              Dim TextForLog As String
91            For i = 1 To sNCols(TheChoices)
92                If IsValidChoice(TheChoices(m_UsersChoice, i)) Then
93                    TextForLog = TextForLog + " > " + Unembellish(CStr(TheChoices(m_UsersChoice, i)))
94                End If
95            Next i
96            TextForLog = "User selected menu item:" + vbLf + Mid$(TextForLog, 4)
97            MessageLogWrite TextForLog
98        End If

99        If ReturnIndex Then
100           ShowCommandBarPopup = m_UsersChoice
101       ElseIf m_UsersChoice = 0 Then
102           ShowCommandBarPopup = "#Cancel!"
103       Else
104           Res = RightMosts(m_UsersChoice, 1)
105           If VarType(Res) = vbString Then
106               Res = Unembellish(CStr(Res))
107           End If
108           ShowCommandBarPopup = Res
109       End If

110       Exit Function
ErrHandler:
111       Throw "#ShowCommandBarPopup(line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function SafeDeleteCommandBar(cb As CommandBar)
1         On Error GoTo ErrHandler
2         cb.Delete
ErrHandler:
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeIsVisible
' Author    : Philip Swannell
' Date      : 26-Apr-2016
' Purpose   : Returns True is some part of a range is visible in the active window
' -----------------------------------------------------------------------------------------------------------------------
Private Function RangeIsVisible(R As Range) As Boolean
          Dim i As Long
1         On Error GoTo ErrHandler
2         If Not ActiveWindow Is Nothing Then
3             For i = 1 To ActiveWindow.Panes.Count
4                 If Not Application.Intersect(R, ActiveWindow.Panes(i).VisibleRange) Is Nothing Then
5                     RangeIsVisible = True
6                     Exit Function
7                 End If
8             Next i
9         End If
10        Exit Function
ErrHandler:
11        Throw "#RangeIsVisible (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : XYCoordinatesOfObjectCentre
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Determine the XY coordinates in the coordinate system used by CommandBar.ShowPopup
'             so as to place the menu with its top left corner at the centre of the passed
'             AnchorObject object. Copes with Range objects, any object that has a TopLeftCell property
'             and CommandButtons on forms. In the case of Ranges and objects with TopLeftCell
'             the method only works correctly if the Range or TopLeftCell is currently visible i.e. in the
'             VisibleRange of the ActiveWindow.
' -----------------------------------------------------------------------------------------------------------------------
Sub XYCoordinatesOfObjectCentre(AnchorObject As Object, ByRef x As Double, ByRef y As Double)
          Dim Parent As Object
          Dim RealAnchorObject As Object
          Dim TheLeft As Double
          Dim TheTop As Double
          Dim X_Nudge As Double
          Dim Y_Nudge As Double

1         On Error GoTo ErrHandler

2         If TypeName(AnchorObject) = "clsPositionInstructions" Then
3             Set RealAnchorObject = AnchorObject.AnchorObject
4             X_Nudge = AnchorObject.X_Nudge
5             Y_Nudge = AnchorObject.Y_Nudge
6         ElseIf Not AnchorObject Is Nothing Then
7             Set RealAnchorObject = AnchorObject
8         End If

9         If Not RealAnchorObject Is Nothing Then
10            If TypeName(RealAnchorObject) = "Range" Then
                  Dim RangeRect As RECT
11                If Not ActiveWindow Is Nothing Then
12                    If RangeIsVisible(RealAnchorObject) Then
13                        RangeRect = GetRangeRect(RealAnchorObject)
14                        x = (RangeRect.Left + RangeRect.Right) / 2
15                        y = (RangeRect.Top + RangeRect.Bottom) / 2
16                    End If
17                End If
18            ElseIf HasTopLeftCell(RealAnchorObject) Then
19                If Not ActiveWindow Is Nothing Then
20                    If RangeIsVisible(RealAnchorObject.TopLeftCell) Then
21                        RangeRect = GetRangeRect(RealAnchorObject.TopLeftCell)
22                        x = (RangeRect.Left) + (RealAnchorObject.Left + RealAnchorObject.Width / 2 - RealAnchorObject.TopLeftCell.Left) * fx
23                        y = (RangeRect.Top) + (RealAnchorObject.Top + RealAnchorObject.Height / 2 - RealAnchorObject.TopLeftCell.Top) * fY
24                    End If
25                End If
26            Else
27                Select Case TypeName(RealAnchorObject)
                      Case "CommandButton", "ListBox", "Label", "TextBox"        'Assume we are dealing with a control on a form
28                        Set Parent = RealAnchorObject
29                        TheTop = Parent.Top
30                        Do While HasParentWithTop(Parent)
31                            Set Parent = Parent.Parent
32                            TheTop = TheTop + Parent.Top
33                        Loop
34                        Set Parent = RealAnchorObject
35                        TheLeft = Parent.Left
36                        Do While HasParentWithTop(Parent)
37                            Set Parent = Parent.Parent
38                            TheLeft = TheLeft + Parent.Left
39                        Loop
40                        x = (TheLeft + RealAnchorObject.Width / 2) * fx
41                        y = (TheTop + (RealAnchorObject.Height + 25) / 2) * fY
42                End Select
43            End If
44        End If

45        x = x + X_Nudge
46        y = y + Y_Nudge

47        Exit Sub
ErrHandler:
48        Throw "#XYCoordinatesOfObjectCentre (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HasTopLeftCell
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Tests if an object has a TopLeftCell property
' -----------------------------------------------------------------------------------------------------------------------
Private Function HasTopLeftCell(obj As Object) As Boolean

1         On Error GoTo ErrHandler
          Dim R As Range

2         Set R = obj.TopLeftCell
3         HasTopLeftCell = True

4         Exit Function
ErrHandler:
5         HasTopLeftCell = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HasParentWithTop
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Tests if an object has a Parent property with a Top property.
' -----------------------------------------------------------------------------------------------------------------------
Private Function HasParentWithTop(obj As Object) As Boolean
          Dim Top As Double

1         On Error GoTo ErrHandler
2         Top = obj.Parent.Top

3         HasParentWithTop = True
4         Exit Function
ErrHandler:
5         HasParentWithTop = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsValidChoice
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Used to test an element of TheChoices to see if we should process it.
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsValidChoice(Choice As Variant) As Boolean
1         Select Case VarType(Choice)
              Case vbBoolean, vbDouble, vbLong, vbInteger, vbCurrency, vbDate, vbSingle
2                 IsValidChoice = True
3             Case vbString
4                 IsValidChoice = Len(Choice) > 0
5         End Select
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowCommandBarOnAction
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : OnAction command of the temporary command bar.
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowCommandBarOnAction()
1         On Error GoTo ErrHandler
2         m_UsersChoice = Application.CommandBars.ActionControl.Tag
3         Exit Sub
ErrHandler:
4         Throw "#ShowCommandBarOnAction(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Unembellish
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : To get BeginGroup lines and underlining of hot keys elements of TheChoices
'             are embellished with ampersand characters and "--" at the start. This method
'             undoes that embellishment and returns the text that the user sees onscreen.
' -----------------------------------------------------------------------------------------------------------------------
Function Unembellish(EmbellishedCaption As String) As Variant

          Dim FoundFirst As Boolean
          Dim FoundSecond As Boolean
          Dim i As Long
          Dim Indicators() As Boolean
          Dim j As Long
          Dim LenResult As Long
          Dim Res As String
          Dim Res2 As String

1         On Error GoTo ErrHandler

2         Res = EmbellishedCaption
3         If Left$(Res, 2) = "--" Then
4             Res = Right$(Res, Len(Res) - 2)
5         End If

6         If InStr(Res, "&") = 0 Then
7             Unembellish = Res
8             Exit Function
9         ElseIf InStr(Res, "&&") = 0 Then
10            Unembellish = Replace(Res, "&", vbNullString)
11            Exit Function
12        End If

          'We need to remove ampersands except two consecutive ampersands. Such pairs become a single ampersand in the result.
13        LenResult = Len(Res)
14        ReDim Indicators(1 To Len(Res))
15        For i = 1 To Len(Res)
16            If Mid$(Res, i, 1) = "&" Then
17                LenResult = LenResult - 1
18                Indicators(i) = True
19                If Not FoundFirst Then
20                    FoundFirst = True
21                Else
22                    FoundSecond = True
23                End If
24                If FoundFirst And FoundSecond Then
25                    Indicators(i - 1) = False
26                    Indicators(i) = False
27                    LenResult = LenResult + 2
28                    FoundFirst = False
29                    FoundSecond = False
30                End If
31            Else
32                FoundFirst = False
33                FoundSecond = False
34            End If
35        Next i

36        Res2 = String(LenResult, " ")
37        For i = 1 To Len(Res)
38            If Not Indicators(i) Then
39                j = j + 1
40                Mid$(Res2, j, 1) = Mid$(Res, i, 1)
41            End If
42        Next

43        Res2 = Replace(Res2, "&&", "&")
44        Unembellish = Res2

45        Exit Function
ErrHandler:
46        Throw "#Unembelish(line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PopulateCommandBar
' Author    : Philip Swannell
' Date      : 11-Nov-2013
' Purpose   : Takes the three arrays TheChoices, EnableFlags and FaceIds and populates
'             the controls of TheCommandBar.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub PopulateCommandBar(ByRef TheCommandBar As Object, _
        ByVal TheChoices As Variant, _
        ByVal EnableFlags As Variant, _
        FaceIDs As Variant)

          Dim AllControls() As Variant
          Dim BeginGroup As Boolean
          Dim ControlType As Long
          Dim FaceID As Variant
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim ParentControl As Object
          Dim thisCaption As String
          Dim ThisControl As CommandBarControl

1         On Error GoTo ErrHandler

2         NC = sNCols(TheChoices)

3         ReDim AllControls(1 To sNRows(TheChoices), 1 To sNCols(TheChoices))

          Dim NR As Long
4         NR = sNRows(TheChoices)

5         For i = 1 To NR
6             For j = 1 To NC
7                 If IsValidChoice(TheChoices(i, j)) Then
8                     thisCaption = TheChoices(i, j)
9                     If Left$(thisCaption, 2) = "--" Then
10                        BeginGroup = True
11                        thisCaption = Right$(thisCaption, Len(thisCaption) - 2)
12                    Else
13                        BeginGroup = False
14                    End If

15                    FaceID = FaceIDs(i, 1)

16                    If j = NC Then
17                        ControlType = msoControlButton
18                    ElseIf Not IsValidChoice(TheChoices(i, j + 1)) Then
19                        ControlType = msoControlButton
20                    Else
21                        ControlType = msoControlPopup
22                    End If
23                    If j = 1 Then
24                        Set ParentControl = TheCommandBar
25                    Else
26                        Set ParentControl = Nothing
27                        For k = i To 1 Step -1
28                            If IsValidChoice(TheChoices(k, j - 1)) Then
29                                Set ParentControl = AllControls(k, j - 1)
30                                If TypeName(ParentControl) <> "CommandBarPopup" Then
31                                    Throw "TheChoices are malformed. Item at row " + CStr(i) + " column " + CStr(j) + " has no parent"
32                                End If
33                                Exit For
34                            End If
35                        Next k
36                        If ParentControl Is Nothing Then
37                            Throw "TheChoices are malformed. Item at row " + CStr(i) + " column " + CStr(j) + " has no parent"
38                        End If
39                    End If

40                    Set ThisControl = ParentControl.Controls.Add(ControlType)
41                    Set AllControls(i, j) = ThisControl

42                    With ThisControl
43                        .caption = thisCaption
44                        .BeginGroup = BeginGroup
45                        If VarType(FaceID) = vbString Then
46                            If Right$(FaceID, 4) = ".bmp" Then
47                                If sFileExists(CStr(FaceID)) Then
48                                    .Picture = LoadPicture(FaceID)
49                                    If EnableFlags(i, 1) Then
50                                        If sFileExists(Left$(FaceID, Len(FaceID) - 4) & "Mask.bmp") Then
51                                            .Mask = LoadPicture(Left$(FaceID, Len(FaceID) - 4) & "Mask.bmp")
52                                        End If
53                                    End If
54                                End If
55                            ElseIf IsInCollection(shCustomIcons.Shapes, CStr(FaceID)) Then
56                                shCustomIcons.Shapes(FaceID).CopyPicture xlScreen, xlBitmap
57                                .Picture = PastePicture(xlBitmap)
58                                If IsInCollection(shCustomIcons.Shapes, CStr(FaceID) & "Mask") Then
59                                    shCustomIcons.Shapes(FaceID & "Mask").CopyPicture xlScreen, xlBitmap
60                                    .Mask = PastePicture(xlBitmap)
61                                End If
62                            Else
63                                .Picture = Application.CommandBars.GetImageMso(FaceID, 16, 16)
64                            End If
65                        ElseIf FaceID > 0 Then
66                            If j = NC Then
67                                .FaceID = FaceID
68                            ElseIf Not IsValidChoice(TheChoices(i, j + 1)) Then
69                                .FaceID = FaceID
70                            End If
71                        End If

72                        If j = NC Then        ' Must be a "leaf of the tree"
73                            .Enabled = EnableFlags(i, 1)
74                        ElseIf Not IsValidChoice(TheChoices(i, j + 1)) Then        'Is a "leaf of the tree"
75                            .Enabled = EnableFlags(i, 1)
76                        Else        'Not a "leaf of the tree", check that at least one descendant is enabled
                              Dim Enable As Boolean
                              Dim x As Long
77                            Enable = False
78                            For x = i To NR
79                                If EnableFlags(x, 1) Then
80                                    Enable = True
81                                    Exit For
82                                End If
83                                If x < NR Then If IsValidChoice(TheChoices(x + 1, j)) Then Exit For
84                            Next x
85                            .Enabled = Enable
86                        End If

87                        If ControlType = msoControlButton Then
88                            .OnAction = ThisWorkbook.Name & "!ShowCommandBarOnAction"
89                            .Tag = i
90                        End If
91                    End With
92                End If
93            Next j
94        Next i

95        Exit Sub
ErrHandler:
96        Throw "#PopulateCommandBar(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sStringWidth
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : Returns the width in points of TextStrings when rendered in the given font. If TextStrings
'             is an array, then the return is an array of the same size.
'
'             1 point is defined as 1/72 of an inch (about 0.35 mm).
' Arguments
' TextStrings: A string or an array of strings.
' FontName  : A font name such as "Calibri", "Arial" or "Garamond".
' FontSize  : The size of the font. E.g. Excel defaults to a size 11 Calibri font.
' FontBold  : TRUE if the font for which sizing is desired is bold.
' FontItalic: TRUE if the font for which sizing is desired is italic.
' -----------------------------------------------------------------------------------------------------------------------
Function sStringWidth(ByVal TextStrings As Variant, _
        FontName As String, FontSize As Long, Optional FontBold As Boolean, Optional FontItalic As Boolean)
Attribute sStringWidth.VB_Description = "Returns the width in points of TextStrings when rendered in the given font. If TextStrings is an array, then the return is an array of the same size.\n\n1 point is defined as 1/72 of an inch (about 0.35 mm)."
Attribute sStringWidth.VB_ProcData.VB_Invoke_Func = " \n25"

          Dim Result() As Variant
          Static frm As frmStringWidth
          Dim Adjustment As Double
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler

          'Mmm PGS 21 Jan 2022 no longer call StringWidth2, original approach seems to work again...

          '  Result2 = sStringWidth2(TextStrings, FontName, FontSize, FontBold, FontItalic)
          '  If Not sIsErrorString(Result2) Then
          '      sStringWidth = Result2
          '      Exit Function
          '  End If

2         If frm Is Nothing Then Set frm = New frmStringWidth

3         Force2DArrayR TextStrings, NR, NC
4         ReDim Result(1 To NR, 1 To NC)

5         With frm.Label1
6             .Font.Name = FontName
7             .Font.Size = FontSize
8             .Font.Bold = FontBold
9             .Font.Italic = FontItalic
10            .caption = "||"
11            .Width = Len("||") * FontSize * 2 + 100
12            .AutoSize = False
13            .AutoSize = True
14            Adjustment = .Width

15            For i = 1 To NR
16                For j = 1 To NC
17                    .caption = "|" & CStr(TextStrings(i, j)) & "|"
18                    .Width = Len(CStr(TextStrings(i, j))) * FontSize * 2 + 100
19                    .AutoSize = False
20                    .AutoSize = True
21                    Result(i, j) = .Width - Adjustment
22                Next j
23            Next i
24            .caption = vbNullString

25        End With
26        sStringWidth = Result

27        Exit Function
ErrHandler:
28        Throw "#sStringWidth (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' ---------------------------------------------------------------------------------------------------------------------------------
' Procedure: sStringWidth2
' Purpose: This is painful. The approach used by method StringWidth, i.e. sizing a label on a hidden form has proved to not work on
'          some PCs, notably the "XPLAIN" PC at Solum that has Windows 10, Office 2016 64 bit. Width numbers come out "wrong".
'          So this is an alternative approach that works for the two cases we use in practice - Calibri 11, and Segoe UI 9
' Author:  Philip Swannell
' Date: 07-Dec-2017
' ---------------------------------------------------------------------------------------------------------------------------------
Private Function sStringWidth2(ByVal TextStrings As Variant, _
        FontName As String, FontSize As Long, Optional FontBold As Boolean, Optional FontItalic As Boolean)

          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long

          Dim Result() As Double
          Dim Widths As Variant

          'Get these numbers from
          'https://d.docs.live.net/4251b448d4115355/Excel Sheets/sStringWidthInvestigation.xlsx

1         On Error GoTo ErrHandler
2         If FontName = "Segoe UI" And FontSize = 9 And FontBold = False And FontItalic = False Then
3             Widths = VBA.Array(9, 6, 6, 6, 6, 6, 6, 2.25, -2.25, 33.75, -2.25, -2.25, -2.25, -2.25, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, _
                  4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 2.25, 2.25, 3.75, 5.25, 4.5, 7.5, 7.5, 2.25, 3, 3, 3.75, 6, 2.25, 3.75, 2.25, 3.75, 4.5, 4.5, _
                  4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 2.25, 2.25, 6, 6, 6, 3.75, 8.25, 6, 5.25, 6, 6, 4.5, 4.5, 6, 6.75, 2.25, 3, _
                  5.25, 4.5, 8.25, 6.75, 6.75, 5.25, 6.75, 5.25, 4.5, 5.25, 6, 5.25, 8.25, 5.25, 5.25, 5.25, 3, 3.75, 3, 6, 3.75, 2.25, 4.5, 5.25, 4.5, _
                  5.25, 4.5, 3, 5.25, 5.25, 2.25, 2.25, 4.5, 2.25, 8.25, 5.25, 5.25, 5.25, 5.25, 3, 3.75, 3, 5.25, 4.5, 6.75, 3.75, 4.5, 3.75, 3, 2.25, _
                  3, 6, 3, 4.5, 6, 2.25, 4.5, 3.75, 6.75, 3.75, 3.75, 3, 11.25, 4.5, 3, 8.25, 6, 5.25, 6, 6, 2.25, 2.25, 3.75, 3.75, 3.75, _
                  4.5, 9, 3, 6.75, 3.75, 3, 8.25, 6, 3.75, 5.25, 2.25, 2.25, 4.5, 4.5, 5.25, 4.5, 2.25, 3.75, 3.75, 8.25, 3.75, 4.5, 6, 0, 8.25, _
                  3.75, 3.75, 6, 3, 3, 2.25, 5.25, 4.5, 2.25, 1.5, 3, 3.75, 4.5, 8.25, 8.25, 8.25, 3.75, 6, 6, 6, 6, 6, 6, 7.5, 6, _
                  4.5, 4.5, 4.5, 4.5, 2.25, 2.25, 2.25, 2.25, 6, 6.75, 6.75, 6.75, 6.75, 6.75, 6.75, 6, 6.75, 6, 6, 6, 6, 5.25, 5.25, 5.25, 4.5, _
                  4.5, 4.5, 4.5, 4.5, 4.5, 7.5, 4.5, 4.5, 4.5, 4.5, 4.5, 2.25, 2.25, 2.25, 2.25, 5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 6, 5.25, 5.25, _
                  5.25, 5.25, 5.25, 4.5, 5.25, 4.5)
4         ElseIf FontName = "Calibri" And FontSize = 11 And FontBold = False And FontItalic = False Then
5             Widths = VBA.Array(11, 6, 6, 6, 6, 6, 6, 2.25, -5.25, 30.75, -5.25, -5.25, -5.25, -5.25, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, _
                  6, 6, 6, 6, 6, 6, 6, 2.25, 3.75, 4.5, 5.25, 5.25, 8.25, 7.5, 2.25, 3.75, 3.75, 5.25, 5.25, 3, 3.75, 3, 4.5, 5.25, 5.25, _
                  5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 3, 3, 5.25, 5.25, 5.25, 5.25, 9.75, 6.75, 6, 6, 6.75, 5.25, 5.25, 6.75, 6.75, 3, 3.75, _
                  6, 4.5, 9, 7.5, 7.5, 6, 7.5, 6, 5.25, 5.25, 6.75, 6.75, 9.75, 6, 5.25, 5.25, 3.75, 4.5, 3.75, 5.25, 5.25, 3, 5.25, 6, 4.5, _
                  6, 6, 3.75, 5.25, 6, 3, 3, 5.25, 3, 9, 6, 6, 6, 6, 3.75, 4.5, 3.75, 6, 5.25, 8.25, 5.25, 5.25, 4.5, 3.75, 5.25, _
                  3.75, 5.25, 6, 5.25, 6, 3, 3.75, 4.5, 7.5, 5.25, 5.25, 4.5, 12, 5.25, 3.75, 9.75, 6, 5.25, 6, 6, 3, 3, 4.5, 4.5, 5.25, _
                  5.25, 10.5, 5.25, 8.25, 4.5, 3.75, 9.75, 6, 4.5, 5.25, 2.25, 3.75, 5.25, 5.25, 5.25, 5.25, 5.25, 5.25, 4.5, 9.75, 4.5, 6, 5.25, 0, 6, _
                  4.5, 3.75, 5.25, 3.75, 3.75, 3, 6, 6.75, 3, 3.75, 3, 4.5, 6, 7.5, 7.5, 7.5, 5.25, 6.75, 6.75, 6.75, 6.75, 6.75, 6.75, 8.25, 6, _
                  5.25, 5.25, 5.25, 5.25, 3, 3, 3, 3, 6.75, 7.5, 7.5, 7.5, 7.5, 7.5, 7.5, 5.25, 7.5, 6.75, 6.75, 6.75, 6.75, 5.25, 6, 6, 5.25, _
                  5.25, 5.25, 5.25, 5.25, 5.25, 9, 4.5, 6, 6, 6, 6, 2.25, 2.25, 2.25, 2.25, 6, 6, 6, 6, 6, 6, 6, 5.25, 6, 6, _
                  6, 6, 6, 5.25, 6, 5.25)
6         Else
7             Throw "Case not handled"
8         End If

          Dim Str As String
          Dim tmp As Double
9         Force2DArrayR TextStrings, NR, NC
10        ReDim Result(1 To NR, 1 To NC)
11        For i = 1 To NR
12            For j = 1 To NC
13                Str = TextStrings(i, j)
14                tmp = 0
15                For k = 1 To Len(Str)
16                    tmp = tmp + Widths(Asc(Mid$(Str, k, 1)))
17                Next k
18                Result(i, j) = tmp
19            Next j
20        Next i
21        sStringWidth2 = Result

22        Exit Function
ErrHandler:
23        sStringWidth2 = "#sStringWidth2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sJustifyArrayOfStrings
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : Returns a column of strings. Each string in the output is the concatenation of the strings
'             in the corresponding row of the input with space characters added such that
'             when rendered in the given font the elements of TextStrings appear correctly
'             aligned.
' Arguments
' TextStrings: An arbitrary array of strings.
' FontName  : A font name such as "Calibri", "Arial" or "Garamond".
' FontSize  : The size of the font in points. E.g. Excel defaults to a size 11 Calibri font.
' ExtraPadding: This string is inserted to increase the spacing between the substrings of the return. If
'             passed as a positive number, then ExtraPadding is set to that number of space
'             characters.
' FontBold  : TRUE if the font in which the output text will be rendered is bold.
' FontItalic: TRUE if the font in which the output text will be rendered is italic.
' -----------------------------------------------------------------------------------------------------------------------
Function sJustifyArrayOfStrings(ByVal TextStrings As Variant, _
        Optional FontName As String = "Calibri", _
        Optional FontSize As Long = 11, _
        Optional ExtraPadding As Variant = "  ", _
        Optional FontBold As Boolean, _
        Optional FontItalic As Boolean, _
        Optional Concatenate As Boolean)
Attribute sJustifyArrayOfStrings.VB_Description = "Returns a column of strings. Each string in the output is the concatenation of the strings in the corresponding row of the input with space characters added such that when rendered in the given font the elements of TextStrings appear correctly aligned."
Attribute sJustifyArrayOfStrings.VB_ProcData.VB_Invoke_Func = " \n25"

          Dim i As Long
          Dim j As Long
          Dim MaxWidthsOfThisColumn As Double
          Dim NC As Long
          Dim NR As Long
          Dim PadString As String
          Dim Result() As Variant
          Dim TheStringWidths As Variant
          Dim WidthOfSpace

1         On Error GoTo ErrHandler

2         Force2DArrayR TextStrings, NR, NC
3         TextStrings = sArrayMakeText(TextStrings)
4         If NC = 1 Then
5             sJustifyArrayOfStrings = TextStrings
6             Exit Function
7         End If
8         WidthOfSpace = sStringWidth(" ", FontName, FontSize, FontBold, FontItalic)(1, 1)
9         TheStringWidths = sStringWidth(sSubArray(TextStrings, 1, 1, , NC - 1), FontName, FontSize, FontBold, FontItalic)
10        If IsMissing(ExtraPadding) Or IsEmpty(ExtraPadding) Then
11            PadString = vbNullString
12        ElseIf IsNumber(ExtraPadding) Then
13            PadString = String(ExtraPadding, " ")
14        ElseIf VarType(ExtraPadding) = vbString Then
15            PadString = ExtraPadding
16        End If

17        ReDim Result(1 To NR, 1 To 1)
18        For i = 1 To NR
19            Result(i, 1) = TextStrings(i, 1)
20        Next
21        For j = 2 To NC
22            MaxWidthsOfThisColumn = 0
23            For i = 1 To NR
24                If TheStringWidths(i, j - 1) > MaxWidthsOfThisColumn Then
25                    MaxWidthsOfThisColumn = TheStringWidths(i, j - 1)
26                End If
27            Next i
28            For i = 1 To NR
29                Result(i, 1) = Result(i, 1) + String((MaxWidthsOfThisColumn - TheStringWidths(i, j - 1)) / WidthOfSpace, " ") + PadString + TextStrings(i, j)
30            Next i
31        Next j
32        If Concatenate Then
33            sJustifyArrayOfStrings = sConcatenateStrings(Result, vbLf)
34        Else
35            sJustifyArrayOfStrings = Result
36        End If

37        Exit Function
ErrHandler:
38        sJustifyArrayOfStrings = "#sJustifyArrayOfStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestsCommentHeight
' Author    : Philip Swannell
' Date      : 03-Aug-2016
' Purpose   : Test harness for method sCommentHeight
' -----------------------------------------------------------------------------------------------------------------------
Sub TestsCommentHeight()
          Dim c As Range
          Dim i As Long
          Dim Res
          Dim TheText As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set ws = Application.Workbooks("Book2").Worksheets(1)
3         Set c = ws.Cells(1, 1)

4         For i = 1 To 200
5             TheText = TheText + String(1 + i Mod 4, Chr$(97 + i Mod 26)) + CStr(i Mod 10) + " "
6         Next

7         c.ClearComments
8         c.AddComment
9         c.Comment.Visible = False
10        c.Comment.text text:=TheText
11        With c.Comment.Shape.TextFrame
12            .Characters.Font.Name = "Calibri"
13            .Characters.Font.Size = 11
14        End With

15        c.Comment.Shape.Width = 200
16        c.Comment.Shape.Height = 400

17        Res = sCommentHeight(TheText, 200)

18        Exit Sub
ErrHandler:
19        SomethingWentWrong "#TestsCommentHeight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCommentHeight
' Author    : Philip Swannell
' Date      : 03-Aug-2016
' Purpose   : Sizing cell comments is difficult, or seems to be. This method estimates the
'             height of a cell comment with width CommentWidth such that TheText will be
'             entirely displayed. Assumes font is Calibri 11.
' -----------------------------------------------------------------------------------------------------------------------
Function sCommentHeight(ByVal TheText As String, CommentWidth As Double)

          Static frm As frmStringWidth
          Static H As Double

1         On Error Resume Next
2         H = frm.Height
3         On Error GoTo ErrHandler
4         If H = 0 Then Set frm = New frmStringWidth

5         With frm.Label1
6             .Font.Name = "Calibri"
7             .Font.Size = 11
8             .Font.Bold = False
9             .Font.Italic = False
10            .caption = TheText + vbLf + "x"        ' one extra line for "safety" - we would rather the return is too large than too small
11            .Width = CommentWidth - 3.5        'margins are narrower in a form label than in a cell comment
12            .AutoSize = False
13            .AutoSize = True
14            sCommentHeight = .Height * 22 / 20.4        ' vertical spacing between lines is smaller in form label vs cell comment
15        End With

16        Exit Function
ErrHandler:
17        Throw "#sCommentHeight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
