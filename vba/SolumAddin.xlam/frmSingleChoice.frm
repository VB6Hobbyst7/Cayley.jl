VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSingleChoice 
   Caption         =   "Select"
   ClientHeight    =   10920
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   5628
   OleObjectBlob   =   "frmSingleChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSingleChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------------------------------------
' Module    : frmSingleChoice
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Implement a dialog to allow a user to choose a value from a list of input
'             strings or numbers. Wrapped by function ShowSingleChoiceDialog.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_clsResizer As clsFormResizer

Public ReturnValue As Variant
Private m_TheChoices As Variant
Private m_TheExplanations As Variant
Private m_FilteredExplanations As Variant
Private m_TheExplanationTitles As Variant
Private m_FilteredExplanationTitles As Variant
Private m_NumRecords As Long
Private m_ShowExplanations As Boolean
Private m_ShowCategories As Boolean
Private m_TheFullTexts As Variant

Const m_WebHelpAvailable = False        '<- PGS 21-Apr-2015. Switch this to True if we write web-based help for the functions such as exists for Martin's TigerLib
Private m_ActiveControlWhenCtrlShiftEnterCaptured As Object
Private m_ActiveControlWhenEnterCaptured As Object
Private m_FullSearchDescription As String
Private m_HelpBrowserMode As Boolean
Private m_RegistrySection As String
Private m_SearchHowMode As Long        '0 = Simple Search, 1 = Starts With, 2 = Regular Expression
Private m_SearchInMode As Long        '0 = Search items only, 1 = Search items and explanations
Private m_SettingsMenu As Variant
Private m_StandardSearchDescription As String
Private m_TheCategories As Variant
Private m_TheCategoriesNoDupes As Variant
Public UseCtrlShiftEnter As Boolean

Private Enum EnmOKContext
    Okc_ClickOK_or_EnterKey = 0        'Mouse click on butOK or Enter key hit irrespective of what control has focus.
    Okc_DoubleClickListBox = 1        'Double Click in LstBxChoices or Windows Right-Click Key hit when LstBxChoices has focus.
    Okc_ClickDownArrow = 2        'Clicked txtbxDownArrow or hit Shift Down when OK but has focus.
    Okc_CtrlShiftEnter = 3        'Ctrl+Shift+Enter hit irrespective of what control has focus.
End Enum
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InitialiseAsHelpBrowser
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Populating and sizing elements on the form, for use as Help Browser, not mere pick-list
' -----------------------------------------------------------------------------------------------------------------------
Sub InitialiseAsHelpBrowser(TheChoices As Variant, _
        TheExplanations, _
        TheExplanationTitles, _
        Title As String, _
        TopText As String, _
        TheCategories As Variant, _
        CategoryLabel As String, _
        RegistryString As String)

1         On Error GoTo ErrHandler

          'Populate module-level variables
2         m_TheChoices = TheChoices
3         m_TheExplanations = TheExplanations
4         m_FilteredExplanations = TheExplanations
5         m_TheExplanationTitles = TheExplanationTitles
6         m_FilteredExplanationTitles = TheExplanationTitles
7         m_NumRecords = sNRows(m_TheChoices)
8         m_TheFullTexts = sArrayConcatenate(m_TheChoices, vbLf, m_TheExplanationTitles, vbLf, m_TheExplanations)

9         m_ShowExplanations = True
10        m_StandardSearchDescription = "&Search topic names only"
11        m_FullSearchDescription = "&Full search in topic names and topic help"
12        m_HelpBrowserMode = True
13        m_ShowCategories = True

14        If RegistryString <> vbNullString Then
15            m_RegistrySection = "ShowSingleChoiceDialog-" & RegistryString
16        End If
17        m_TheCategories = TheCategories
18        m_TheCategoriesNoDupes = sArrayStack("All", sRemoveDuplicates(TheCategories, True))

          'Lay out the controls
19        PositionControlsAsHelpBrowser Me, TheChoices, TheExplanations, TheExplanationTitles, Title, TopText, TheCategories, CategoryLabel, RegistryString, m_RegistrySection, m_TheCategoriesNoDupes, m_SearchInMode, m_SearchHowMode

20        UpdateNumRecords Me, m_ShowCategories, m_NumRecords, m_SearchInMode, m_SearchHowMode, m_ShowExplanations, m_StandardSearchDescription, m_FullSearchDescription
21        UpdateExplanation Me, m_ShowExplanations, m_HelpBrowserMode, m_WebHelpAvailable, m_FilteredExplanations, m_FilteredExplanationTitles
22        butOK_EnableOrDisable

23        SetUpResizerHelpBrowserMode Me, m_clsResizer

24        Exit Sub
ErrHandler:
25        Throw "#InitialiseAsHelpBrowser (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MinWidthForComboBox
' Author    : Philip
' Date      : 11-Oct-2017
' Purpose   : ComboBoxes and ListBoxes don't have an autosize method, which is annoying.
'             Use the lblAutoSize label to roll our own
' -----------------------------------------------------------------------------------------------------------------------
Function MinWidthForComboBox(cmb As Object, Optional ByVal List As Variant)
1         On Error GoTo ErrHandler
2         If IsMissing(List) Then
3             List = cmb.List
4         End If
5         lblAutoSize.Font.Name = cmb.Font.Name
6         lblAutoSize.Font.Size = cmb.Font.Size
7         lblAutoSize.caption = vbLf + sConcatenateStrings(List, vbLf)
8         lblAutoSize.Width = 10000
9         lblAutoSize.AutoSize = False
10        lblAutoSize.AutoSize = True
11        MinWidthForComboBox = lblAutoSize.Width + 25
12        lblAutoSize.caption = vbNullString
13        lblAutoSize.Visible = False
14        Exit Function
ErrHandler:
15        Throw "#MinWidthForComboBox (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ChoiceIsClear
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Encapsulate decision if the user has uniquely defined a return value
' -----------------------------------------------------------------------------------------------------------------------
Private Function ChoiceIsClear() As Boolean
1         On Error GoTo ErrHandler
          Dim ChosenFunction As String

2         If Not IsNull(LstBxChoices.Value) Then
3             ChoiceIsClear = True
4             ChosenFunction = LstBxChoices.Value
5         ElseIf LstBxChoices.ListCount = 1 Then
6             ChoiceIsClear = True
7             ChosenFunction = LstBxChoices.List(0)
8         End If

          'Cope with topics that aren't functions e.g. "Ctrl Shift R" - may need a better way of _
           distinguishing between topics that are functions and topics that are not, but for the time being this will work.
9         If m_HelpBrowserMode Then
10            If InStr(ChosenFunction, " ") > 0 Then
11                ChoiceIsClear = False
12            End If
13        End If

14        Exit Function
ErrHandler:
15        Throw "#ChoiceIsClear (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub butCancel_Click()

1         On Error GoTo ErrHandler
2         Me.ReturnValue = Empty
3         HideForm Me

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butCancel_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClickedOK
' Author    : Philip Swannell
' Date      : 12-Nov-2013
' Purpose   : Called in a variety of contexts (see enumeration EnmOKContext). If there are problems which will cause
'             the formula insertion to not work or if we need confirmation from the user for overwriting
'             then it looks less jarring to post message boxes without dismissing the dialog.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ClickedOK(Context As EnmOKContext)

          Dim EnableFlags As Variant
          Dim FaceIDs As Variant
          Dim Key1 As String
          Dim Key2 As String
          Dim Key3 As String
          Dim Key4 As String
          Dim Option1 As String
          Dim Option2 As String
          Dim Option3 As String
          Dim Option4 As String
          Dim PI As clsPositionInstructions
          Dim PopupRes As Long
          Dim TheChoices As Variant
          Dim TheKeys As Variant
          Dim UsersChoice As String
          Dim Y_Nudge As Double

1         On Error GoTo ErrHandler

2         If Context = Okc_CtrlShiftEnter Then
3             If Not m_HelpBrowserMode Then
4                 GoTo EarlyExit
5             ElseIf Not ChoiceIsClear() Then
6                 GoTo EarlyExit
7             End If
8         End If

9         If Context = Okc_ClickOK_or_EnterKey Then
10            If Not ChoiceIsClear Then
11                GoTo EarlyExit
12            End If
13        End If

14        If Not IsNull(LstBxChoices.Value) Then
15            UsersChoice = LstBxChoices.Value
16        ElseIf LstBxChoices.ListCount = 1 Then
17            UsersChoice = LstBxChoices.List(0)
18        ElseIf Context <> Okc_ClickDownArrow Then
19            GoTo EarlyExit
20        End If

21        If m_HelpBrowserMode Then

22            If Context = Okc_ClickOK_or_EnterKey Then
23                PopupRes = 2
24            ElseIf Context = Okc_CtrlShiftEnter Then
25                PopupRes = 1
26            Else
27                Option1 = "Insert &Array formula {=" + UsersChoice + "(...)}"
28                Key1 = "Ctrl Shift Enter"
29                Option2 = "Insert &non-array formula =" + UsersChoice + "(...)"
30                Key2 = "Enter"
31                Option3 = "--&Help on " + UsersChoice
32                Key3 = "F1"
33                Option4 = "--Search &Options"
34                Key4 = "F2"

35                If Context = Okc_DoubleClickListBox Then
36                    Set PI = New clsPositionInstructions
37                    Set PI.AnchorObject = LstBxChoices
38                    Y_Nudge = -LstBxChoices.Height / 2 + 11
39                    Y_Nudge = Y_Nudge + (LstBxChoices.ListIndex - LstBxChoices.TopIndex) * 9.75
40                    Y_Nudge = Y_Nudge * fY()
41                    PI.Y_Nudge = Y_Nudge

42                    TheChoices = sArrayStack(Option1, Option2, Option3)
43                    TheKeys = sArrayStack(Key1, Key2, Key3)
44                    FaceIDs = sArrayStack(2637, 2474, 49)
45                    EnableFlags = sArrayStack(True, True, True)
46                ElseIf Context = Okc_ClickDownArrow Then
47                    Set PI = New clsPositionInstructions
48                    Set PI.AnchorObject = butOK
49                    PI.Y_Nudge = (butOK.Height / 2 + 6) * fY()
50                    PI.X_Nudge = (-butOK.Width / 2 + 3) * fx()

51                    TheChoices = sArrayStack(Option1, Option2, Option3, Option4)
52                    TheKeys = sArrayStack(Key1, Key2, Key3, Key4)
53                    FaceIDs = sArrayStack(2637, 2474, 49, 1714)
54                    EnableFlags = sArrayStack(True, True, True, True)
55                End If
56                If Not IsExcelReadyForFormula(True) Then        'Morph the menu for when Excel is not ready - e.g. no active worksheet
57                    EnableFlags(1, 1) = False: EnableFlags(2, 1) = False
58                    FaceIDs(1, 1) = 0: FaceIDs(2, 1) = 0
59                End If
60                If UsersChoice = vbNullString Then
61                    EnableFlags(1, 1) = False: EnableFlags(2, 1) = False: EnableFlags(3, 1) = False
62                    FaceIDs(1, 1) = 0: FaceIDs(2, 1) = 0        'Set these FacIDs to 0 since my home-made icons render poorly in not-enabled mode
63                    TheChoices(1, 1) = "Insert Array Formula"
64                    TheChoices(2, 1) = "Insert Non Array Formula"
65                    TheChoices(3, 1) = "Help"
66                End If

67                If Not m_WebHelpAvailable Then
                      Dim HChooseVector As Variant
68                    If sNRows(TheChoices) = 4 Then
69                        HChooseVector = sArrayStack(True, True, False, True)
70                    Else
71                        HChooseVector = sArrayStack(True, True, False)
72                    End If
73                    TheChoices = sMChoose(TheChoices, HChooseVector)
74                    TheKeys = sMChoose(TheKeys, HChooseVector)
75                    FaceIDs = sMChoose(FaceIDs, HChooseVector)
76                    EnableFlags = sMChoose(EnableFlags, HChooseVector)
77                End If

78                TheChoices = sConcatenateLabelsAndKeyDescriptions(TheChoices, TheKeys, 8)

                  'The menu that appears states that F1 takes the user to TigerHelp, which it does, _
                   except when the menu itself is visible. I tried using Application.OnKey to fix that _
                   small problem, but to no avail.
79                PopupRes = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , PI, True)
80                UnHighlight butOK

81                If PopupRes = 0 Then GoTo EarlyExit
82            End If

83            If PopupRes = 1 Or PopupRes = 2 Then
84                If Not IsExcelReadyForFormula(False, PopupRes = 1) Then GoTo EarlyExit
85                If Not ConfirmInsertFunction(PopupRes = 1, UsersChoice) Then GoTo EarlyExit
86            End If

87            If PopupRes = 0 Then
88                GoTo EarlyExit
89            ElseIf PopupRes = 1 Then
90                Me.UseCtrlShiftEnter = True
91            ElseIf PopupRes = 2 Then
92                Me.UseCtrlShiftEnter = False
93            ElseIf m_WebHelpAvailable Then
94                If PopupRes = 3 Then
95                    ShowHelpForFunction UsersChoice
96                    GoTo EarlyExit
97                ElseIf PopupRes = 4 Then
98                    lblSearchOptions_Click
99                    GoTo EarlyExit
100               End If
101           ElseIf PopupRes = 3 Then
102               lblSearchOptions_Click
103               GoTo EarlyExit
104           End If
105       End If

106       Me.ReturnValue = UsersChoice

107       HideForm Me

108       Exit Sub

EarlyExit:
109       Exit Sub
ErrHandler:
110       Throw "#ClickedOK (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sConcatenateLabelsAndKeyDescriptions
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : For use in constructing the TheChoices argument to ShowCommandBarPopup
'             so that if we want to mention a keyboard equivalent to each command in the menu
'             then those keyboard equivalents are right justified in the menu.
' -----------------------------------------------------------------------------------------------------------------------
Private Function sConcatenateLabelsAndKeyDescriptions(Labels As Variant, _
        KeyDescriptions As Variant, _
        Optional NumExtraSpaces As Long = 5)

          Dim CleanedLabels As Variant
          Dim i As Long
          Dim MaxWidth As Double
          Dim NR As Long
          Dim Result() As Variant
          Dim TheStringWidths As Variant
          Dim WidthOfSpace
          Const FontName = "Segoe UI"
          Const FontSize = 9

1         On Error GoTo ErrHandler

2         Force2DArrayR Labels, NR
3         CleanedLabels = Labels
4         Force2DArrayR KeyDescriptions

5         WidthOfSpace = sStringWidth(" ", FontName, FontSize)(1, 1)
6         For i = 1 To NR
7             CleanedLabels(i, 1) = Unembellish(CStr(Labels(i, 1)))
8         Next i

9         TheStringWidths = sStringWidth(CleanedLabels, FontName, FontSize)

10        ReDim Result(1 To NR, 1 To 1)
11        Result = Labels
12        MaxWidth = 0
13        For i = 1 To NR
14            If TheStringWidths(i, 1) > MaxWidth Then
15                MaxWidth = TheStringWidths(i, 1)
16            End If
17        Next i
18        For i = 1 To NR
19            Result(i, 1) = Result(i, 1) + String((MaxWidth - TheStringWidths(i, 1)) / WidthOfSpace + NumExtraSpaces, " ") + KeyDescriptions(i, 1)
20        Next i
21        sConcatenateLabelsAndKeyDescriptions = Result

22        Exit Function
ErrHandler:
23        Throw "#sConcatenateLabelsAndKeyDescriptions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub butCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         CaptureKeys KeyCode, Shift

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butCancel_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         Highlight butCancel
3         UnHighlight butOK

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butCancel_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butOK_Click()
1         On Error GoTo ErrHandler

2         If Not m_ActiveControlWhenEnterCaptured Is Nothing Then
3             If Not m_ActiveControlWhenCtrlShiftEnterCaptured Is Nothing Then
4                 Throw "Assertion Failure: Logic error: variables m_ActiveControlWhenEnterCaptured and m_ActiveControlWhenCtrlShiftEnterCaptured both exist simultaneously"
5             End If
6         End If

7         If Not m_ActiveControlWhenEnterCaptured Is Nothing Then
8             m_ActiveControlWhenEnterCaptured.SetFocus
9             Set m_ActiveControlWhenEnterCaptured = Nothing
10            Exit Sub
11        End If

12        If Not m_ActiveControlWhenCtrlShiftEnterCaptured Is Nothing Then
13            m_ActiveControlWhenCtrlShiftEnterCaptured.SetFocus
14            Set m_ActiveControlWhenCtrlShiftEnterCaptured = Nothing
15            Exit Sub
16        End If

17        ClickedOK Okc_ClickOK_or_EnterKey

18        Exit Sub
ErrHandler:
19        SomethingWentWrong "#butOK_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 40 And Shift = 1 Then        ' Repond to Shift Down
3             If m_HelpBrowserMode Then
4                 If ChoiceIsClear() Then
5                     ClickedOK Okc_ClickDownArrow
6                 End If
7             End If
8         ElseIf KeyCode = 38 And Shift = 0 Then        'Up key - otherwise the upkey causes txtBxExplanations to _
                                                         get focus and that is strange - it doesn't look as though _
                                                         it should be able to receive focus.
9             LstBxChoices.SetFocus
10        End If

11        CaptureKeys KeyCode, Shift

12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#butOK_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         Highlight butOK
3         UnHighlight butCancel

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butOK_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Highlight
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : Form dialogs are a bit old-school, for example buttons don't get highlighted
'             as the mouse is moved over them. This method and Unhighlight make that happen.
' -----------------------------------------------------------------------------------------------------------------------
Sub Highlight(o As Object)
          'Other colors tried: RGB(222, 242, 252),&H80000005,&H80000014&
          Dim TheCol As Double
1         On Error GoTo ErrHandler

2         If o Is butOK Then
3             If Not ChoiceIsClear Then
4                 Exit Sub
5             End If
6         End If

7         TheCol = &H80000014
8         If o.BackColor <> TheCol Then
9             o.BackColor = TheCol
10        End If
11        If o Is butOK Then
12            If m_HelpBrowserMode Then
13                If lblDivider.BorderStyle <> fmBorderStyleSingle Then
14                    lblDivider.BorderStyle = fmBorderStyleSingle
15                End If
16            End If
17        End If

18        Exit Sub
ErrHandler:
19        Throw "#Highlight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Sub UnHighlight(o As Object)
1         On Error GoTo ErrHandler

2         If o.BackColor <> &H8000000F Then
3             o.BackColor = &H8000000F
4         End If
5         If o Is butOK Then
6             If m_HelpBrowserMode Then
7                 If lblDivider.BorderStyle <> fmBorderStyleNone Then
8                     lblDivider.BorderStyle = fmBorderStyleNone
9                 End If
10            End If
11        End If

12        Exit Sub
ErrHandler:
13        Throw "#UnHighlight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub CmbBxCategories_Change()
1         On Error GoTo ErrHandler

2         TxtBxFilter_Change

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#CmbBxCategories_Change (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub CmbBxCategories_Enter()
1         On Error GoTo ErrHandler

2         CmbBxCategories.DropDown

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#CmbBxCategories_Enter (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub CmbBxCategories_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
2         CaptureKeys KeyCode, Shift
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#CmbBxCategories_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AssistConstructFilter
' Author    : Philip Swannell
' Date      : 03-May-2016
' Purpose   : Pop up a dialog to help with constructing a regular expression for filtering
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AssistConstructFilter()
          Dim ActionText As String
          Dim AttributeName As String
          Dim CustomRegEx
          Dim RegKey As String
          Dim Title As String
          Dim WithMRU As Boolean
1         On Error GoTo ErrHandler
2         m_SearchHowMode = 2
3         If m_SearchHowMode = 2 Then
4             If m_HelpBrowserMode Then
5                 Title = "Filter functions"
6                 If m_SearchInMode = 0 Then
7                     AttributeName = "The topic name"
8                 Else
9                     AttributeName = "The topic name or topic help"
10                End If
11                ActionText = "Show topics where"
12                RegKey = "HelpBrowser"
13                WithMRU = False
14            Else
15                Title = "Filter items"
16                AttributeName = "The item"
17                ActionText = "Show items where"
18                WithMRU = False
19            End If

20            CustomRegEx = ShowRegularExpressionDialog(TxtBxFilter.Value, AttributeName, , Me, Title, ActionText, WithMRU, RegKey, False)
21            If CustomRegEx <> "#User Cancel!" Then
22                Me.TxtBxFilter.Value = CustomRegEx
23            End If
24        End If
25        Exit Sub
ErrHandler:
26        Throw "#AssistConstructFilter (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub ShowHelp()
          Dim HelpText As String
1         On Error GoTo ErrHandler
2         HelpText = "Simple match" & vbLf & _
              "Items displayed are those that contain the text you type (case insensitive)." & vbLf & _
              vbNullString & vbLf & _
              "'Starts with' match" & vbLf & _
              "Items displayed are those that start with the text you type (case insensitive)."

3         HelpText = HelpText & vbLf & _
              vbNullString & vbLf & _
              "Regular Expression match" & vbLf & _
              "A powerful way to search for patterns in text."

4         HelpText = HelpText + " Double-click in '" + Replace(Replace(lblTopText.caption, " ", Chr$(28)), ":", vbNullString) + _
              "' to make your own Regular Expression or see the guide below:" + vbLf + vbLf

5         HelpText = HelpText + "Character classes" & vbLf & _
              ".                          any character except newline" & vbLf & _
              "\w \d \s           word, digit, whitespace" & vbLf & _
              "\W \D \S          not word, not digit, not whitespace" & vbLf & _
              "[abc]                 any of a, b, or c" & vbLf & _
              "[^abc]               not a, b, or c" & vbLf & _
              "[a-g]                  character between a & g" & vbLf & vbLf

6         HelpText = HelpText + "Anchors" & vbLf & _
              "^abc$                start / end of the string" & vbLf & _
              "\b                       word boundary" & vbLf & vbLf & _
              "Escaped characters" & vbLf & _
              "\. \* \\              escaped special characters" & vbLf & _
              "\t \n \r              tab, linefeed, carriage return" & vbLf & vbLf _
      
7             HelpText = HelpText + "Groups and Lookarounds" & vbLf & _
                  "(abc)                 capture group" & vbLf & _
                  "\1                       backreference to group #1" & vbLf & _
                  "(?:abc)              non-capturing group" & vbLf & _
                  "(?=abc)             positive lookahead" & vbLf & _
                  "(?!abc)             negative lookahead" & vbLf & vbLf & _
                  "Quantifiers and Alternation" & vbLf & _
                  "a* a+ a?            0 or more, 1 or more, 0 or 1" & vbLf & _
                  "a{5} a{2,}         exactly five, two or more" & vbLf & _
                  "a{1,3}                between one & three" & vbLf & _
                  "a+? a{2,}?        match as few as possible" & vbLf & _
                  "ab|cd                match ab or cd" & vbLf & vbLf

8             HelpText = HelpText + "For more help follow the links below:" + vbLf + _
                  "http://www.regular-expressions.info/" + vbLf + _
                  "https://en.wikipedia.org/wiki/Regular_expression"

9             MsgBoxPlus HelpText, vbOKOnly + vbInformation, "Help on Matching", , , , , 350, , , , , Me
10            Exit Sub
ErrHandler:
11            Throw "#ShowHelp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : lblSearchOptions_Click
' Author    : Philip Swannell
' Date      : 12-Nov-2013
' Purpose   : Popup the menu for changing search options. Also hooked to F2.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub lblSearchOptions_Click()
          Dim EnableFlags As Variant
          Dim PI As clsPositionInstructions
          Dim Res As Variant
          Dim TheChoices As Variant
          Dim TheFaceIDs As Variant

1         On Error GoTo ErrHandler

          Const chSimple = "--Si&mple match"
          Const chStartsWith = "'Starts &with' match"
          Const chRegExp = "&Regular Expression match"
          Const chRegExCustom = "Make a Regular Expression...                            (Double-click)"
          Const chHelp = "--&Help on matching"

2         If m_ShowExplanations Then
3             TheChoices = sArrayStack(m_StandardSearchDescription, m_FullSearchDescription, chSimple, chStartsWith, chRegExp, _
                  "--" & chRegExCustom, chHelp)
4         Else
5             TheChoices = sArrayStack(chSimple, chStartsWith, chRegExp, _
                  "--" & chRegExCustom, chHelp)
6         End If

7         TheFaceIDs = sReshape(0, sNRows(TheChoices), 1)
8         If m_ShowExplanations Then
9             TheFaceIDs(m_SearchInMode + 1, 1) = 1087
10            TheFaceIDs(m_SearchHowMode + 3, 1) = 1087
11        Else
12            TheFaceIDs(m_SearchHowMode + 1, 1) = 1087
13        End If
14        TheFaceIDs(sMatch(chHelp, TheChoices), 1) = 49

15        EnableFlags = sReshape(True, sNRows(TheChoices), 1)

16        CreatePositionInstructions PI, lblSearchOptions, _
              (-lblSearchOptions.Width / 2 + 3) * fY, _
              (lblSearchOptions.Height / 2 + 6) * fY

17        If Not IsEmpty(m_SettingsMenu) Then
              'm_Settings defines further menu elements, has three columns: FaceID, Choice, Macro
18            TheFaceIDs = sArrayStack(TheFaceIDs, sSubArray(m_SettingsMenu, 1, 1, , 1))
19            TheChoices = sArrayStack(TheChoices, sSubArray(m_SettingsMenu, 1, 2, , 1))
20            EnableFlags = sArrayStack(EnableFlags, sReshape(True, sNRows(m_SettingsMenu), 1))
21        End If

22        Res = ShowCommandBarPopup(TheChoices, TheFaceIDs, EnableFlags, , PI, False)

          'Code in this section must be in synch with corresponding calls _
           to GetSettings in method Initialise.
23        If Res = "#Cancel!" Then
24            Exit Sub
25        ElseIf Res = Unembellish(m_StandardSearchDescription) Then
26            m_SearchInMode = 0
27            If m_RegistrySection <> vbNullString Then
28                SaveSetting gAddinName, m_RegistrySection, "SearchInMode", "ItemsOnly"
29            End If
30        ElseIf Res = Unembellish(m_FullSearchDescription) Then
31            m_SearchInMode = 1
32            If m_RegistrySection <> vbNullString Then
33                SaveSetting gAddinName, m_RegistrySection, "SearchInMode", "ItemsAndDescriptions"
34            End If
35        ElseIf Res = Unembellish(chSimple) Then
36            m_SearchHowMode = 0
37            If m_RegistrySection <> vbNullString Then
38                SaveSetting gAddinName, m_RegistrySection, "SearchHowMode", "Simple"
39            End If
40        ElseIf Res = Unembellish(chStartsWith) Then
41            m_SearchHowMode = 1
42            If m_RegistrySection <> vbNullString Then
43                SaveSetting gAddinName, m_RegistrySection, "SearchHowMode", "StartsWith"
44            End If
45        ElseIf Res = Unembellish(chRegExp) Then
46            m_SearchHowMode = 2
47            If m_RegistrySection <> vbNullString Then
48                SaveSetting gAddinName, m_RegistrySection, "SearchHowMode", "RegExp"
49            End If
50        ElseIf Res = Unembellish(chRegExCustom) Then
51            AssistConstructFilter

52        ElseIf Res = Unembellish(chHelp) Then
53            ShowHelp
54        ElseIf Not IsEmpty(m_SettingsMenu) Then
              Dim i As Long
55            For i = 1 To sNRows(m_SettingsMenu)
56                If Res = Unembellish(CStr(m_SettingsMenu(i, 2))) Then
57                    If m_SettingsMenu(i, 3) = "ClearHistory" Then
58                        ClearHistory
59                    Else
60                        Application.Run m_SettingsMenu(i, 3)
61                    End If
62                    Exit For
63                End If
64            Next i
65        End If

66        TxtBxFilter_Change

67        Exit Sub
ErrHandler:
68        SomethingWentWrong "#lblSearchOptions_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearHistory
' Author    : Philip
' Date      : 13-Oct-2017
' Purpose   : Shameful special case coding - only gets called when the RObjectViewer
'             workbook is in use, since in that case an extra menu item "Clear History"
'             appears in the menu that's attached to lblSearchOptions
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ClearHistory()
1         On Error Resume Next
2         DeleteSetting gAddinName, "RObjectViewer"
3         On Error GoTo ErrHandler

4         m_TheChoices = sMChoose(m_TheChoices, sArrayNot(sArrayEquals(m_TheCategories, "History")))
5         m_NumRecords = sNRows(m_TheChoices)
6         m_TheCategories = sMChoose(m_TheCategories, sArrayNot(sArrayEquals(m_TheCategories, "History")))
7         m_TheCategoriesNoDupes = sMChoose(m_TheCategoriesNoDupes, sArrayNot(sArrayEquals(m_TheCategoriesNoDupes, "History")))
8         CmbBxCategories.List = m_TheCategoriesNoDupes
9         TxtBxFilter_Change
10        Exit Sub
ErrHandler:
11        Throw "#ClearHistory (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub lblHelp_Click()
          Dim UsersChoice As String
1         On Error GoTo ErrHandler

2         If Not IsNull(LstBxChoices.Value) Then
3             UsersChoice = LstBxChoices.Value
4         ElseIf LstBxChoices.ListCount = 1 Then
5             UsersChoice = LstBxChoices.List(0)
6         Else
7             Exit Sub
8         End If
9         ShowHelpForFunction UsersChoice

10        Exit Sub
ErrHandler:
11        SomethingWentWrong "#lblHelp_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub LstBxChoices_Change()
1         On Error GoTo ErrHandler

2         butOK_EnableOrDisable
3         UpdateExplanation Me, m_ShowExplanations, m_HelpBrowserMode, m_WebHelpAvailable, m_FilteredExplanations, m_FilteredExplanationTitles

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#LstBxChoices_Change (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub LstBxChoices_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler

2         ClickedOK Okc_DoubleClickListBox

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#LstBxChoices_DblClick (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub LstBxChoices_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         UnHighlight butOK
3         UnHighlight butCancel

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#LstBxChoices_MouseDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub txtbxDownArrow_Enter()
1         On Error GoTo ErrHandler

2         butOK.SetFocus
3         ClickedOK Okc_ClickDownArrow

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#txtbxDownArrow_Enter (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub txtbxDownArrow_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         CaptureKeys KeyCode, Shift

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#txtbxDownArrow_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub txtbxDownArrow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         Highlight butOK
3         UnHighlight butCancel

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#txtbxDownArrow_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub TxtBxExplanations_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If (KeyCode = 67 And Shift = 2) Or (KeyCode = 45 And Shift = 2) Then        'Ctrl C or Ctrl Insert
3             If TxtBxExplanations.SelLength > 0 Then
4                 CopyStringToClipboard TxtBxExplanations.SelText
5             End If
6         End If

7         CaptureKeys KeyCode, Shift

8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#TxtBxExplanations_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub TxtBxExplanations_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlight butOK
3         UnHighlight butCancel
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TxtBxExplanations_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TxtBxExplanations_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         TextBoxMouseEventB TxtBxExplanations, Button, Shift, x, y, False, m_TheChoices, m_HelpBrowserMode, LstBxChoices, TxtBxFilter
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TxtBxExplanations_MouseUp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TxtBxFilter_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         Cancel.Value = True
3         AssistConstructFilter
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TxtBxFilter_DblClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         UnHighlight butCancel
3         UnHighlight butOK

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub UserForm_Resize()
1         On Error GoTo ErrHandler

2         KickScrollBars Me, m_ShowExplanations

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#UserForm_Resize (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LstBxChoices_KeyDown
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : When the user has scrolled to the top of the list we want a further hit of
'             the up key to take focus to the TxtBxFilter.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub LstBxChoices_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler

2         If KeyCode = 38 Then        'Upkey
3             If LstBxChoices.ListIndex = 0 Then        'we are at the top of the list
4                 TxtBxFilter.SetFocus
5                 LstBxChoices.Enabled = False
6                 LstBxChoices.ListIndex = -1
7                 LstBxChoices.Enabled = True
8                 LstBxChoices.Selected(0) = False
9             End If
10        ElseIf KeyCode = 93 Then        'Right click key - to the left of the right-hand _
                                           Ctrl key on my keyboard. Simulate a Right-Click event , _
                                           though like all Form controls this control doesn't _
                                           even have a right-click event.
11            If m_HelpBrowserMode Then
12                ClickedOK Okc_DoubleClickListBox
13            End If
14        End If

15        CaptureKeys KeyCode, Shift

16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#LstBxChoices_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CaptureKeys
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : Make a couple of key combinations work no matter what control has focus.
'             This method is called from the KeyDown event of all controls that
'             can have focus - OK since not many controls can have focus. If there were
'             a great many such controls a possible technique is described at
'             http://www.xtremevbtalk.com/showthread.php?t=305400
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CaptureKeys(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 13 And Shift = 0 Then        ' Enter key
3             Set m_ActiveControlWhenEnterCaptured = Me.ActiveControl
4             ClickedOK Okc_ClickOK_or_EnterKey
5         End If

6         If m_HelpBrowserMode Then
7             If KeyCode = 13 Then        'Enter
8                 If Shift = 3 Then        'Ctrl Shift
9                     If ChoiceIsClear() Then
10                        Set m_ActiveControlWhenCtrlShiftEnterCaptured = Me.ActiveControl
11                        ClickedOK Okc_CtrlShiftEnter
12                    End If
13                End If
14            End If
15        End If
16        If m_HelpBrowserMode Then
17            If KeyCode = 112 Then        'F1
18                If ChoiceIsClear Then
19                    lblHelp_Click
20                End If
21            End If
22        End If
23        If m_ShowExplanations Then
24            If KeyCode = 113 Then        'F2
25                lblSearchOptions_Click
26            End If
27        End If

28        Exit Sub
ErrHandler:
29        Throw "#CaptureKeys (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TxtBxFilter_Change
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Change the contents of the list box as user types into the edit box, also
'             called from the Change event of CmbBxCategories.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TxtBxFilter_Change()
          Dim ChooseVector As Variant
          Dim FilterByCategory As Boolean
          Dim i As Long
          Dim Res As Variant

1         On Error GoTo ErrHandler

2         If m_ShowCategories Then
3             If CmbBxCategories.Value <> "All" Then
4                 FilterByCategory = True
5             End If
6         End If

7         If TxtBxFilter.Value = vbNullString And Not FilterByCategory Then
8             m_FilteredExplanations = m_TheExplanations
9             m_FilteredExplanationTitles = m_TheExplanationTitles
10            LstBxChoices.List = m_TheChoices
11        Else

12            If m_SearchHowMode = 0 Then        'Simple Match
13                If TxtBxFilter.Value = vbNullString Then
14                    ChooseVector = sReshape(True, m_NumRecords, 1)
15                Else
16                    If m_SearchInMode = 0 Then
17                        ChooseVector = sArrayFind(TxtBxFilter.Value, m_TheChoices)
18                    Else
19                        ChooseVector = sArrayFind(TxtBxFilter.Value, m_TheFullTexts)
20                    End If
21                End If

22            ElseIf m_SearchHowMode = 1 Then        'Starts With Match
                  'Don't bother to look at the explanations in StartsWith mode
23                ChooseVector = sArrayEquals(sArrayLeft(m_TheChoices, Len(TxtBxFilter.Value)), TxtBxFilter.Value, False)
24            ElseIf m_SearchHowMode = 2 Then        'Regular Expression mode
                  Dim RegularExpression As String
25                RegularExpression = TxtBxFilter.Value
26                If m_SearchInMode = 0 Then
27                    ChooseVector = sArrayEquals(True, sIsRegMatch(RegularExpression, m_TheChoices, False))
28                Else
29                    ChooseVector = sArrayEquals(True, sIsRegMatch(RegularExpression, m_TheFullTexts, False))
30                End If
31            End If

32            If FilterByCategory Then
33                For i = 1 To m_NumRecords
34                    ChooseVector(i, 1) = ChooseVector(i, 1) And sEquals(m_TheCategories(i, 1), CmbBxCategories.Value)
35                Next i
36            End If

37            If VarType(ChooseVector) = vbString Then
38                m_FilteredExplanations = Empty
39                m_FilteredExplanationTitles = Empty
40                LstBxChoices.Clear
41            Else
42                If sArrayCount(ChooseVector) = 0 Then
43                    m_FilteredExplanations = Empty
44                    m_FilteredExplanationTitles = Empty
45                    LstBxChoices.Clear
46                Else
47                    Res = sMChoose(m_TheChoices, ChooseVector)
48                    m_FilteredExplanations = sMChoose(m_TheExplanations, ChooseVector)
49                    m_FilteredExplanationTitles = sMChoose(m_TheExplanationTitles, ChooseVector)
50                    Force2DArray m_FilteredExplanations
51                    Force2DArray m_FilteredExplanationTitles
52                    Force2DArray Res
53                    LstBxChoices.List = Res
54                End If
55            End If
56        End If

57        UpdateNumRecords Me, m_ShowCategories, m_NumRecords, m_SearchInMode, m_SearchHowMode, m_ShowExplanations, m_StandardSearchDescription, m_FullSearchDescription
58        UpdateExplanation Me, m_ShowExplanations, m_HelpBrowserMode, m_WebHelpAvailable, m_FilteredExplanations, m_FilteredExplanationTitles
59        butOK_EnableOrDisable

60        Exit Sub
ErrHandler:
61        SomethingWentWrong "#TxtBxFilter_Change (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Sub butOK_EnableOrDisable()
          'We don't disable the control because _
           disabling the Default control has consequences for what happens when _
           the user hits Enter in other controls.  Hitting Enter becomes the same _
           as hitting Tab, which is not the user's intention (at least not mine). _
           So instead we just make the button look disabled.
1         On Error GoTo ErrHandler
2         If ChoiceIsClear Then
3             butOK.ForeColor = &H80000012
4             txtbxDownArrow.ForeColor = &H80000012
5             butOK.TakeFocusOnClick = True
6             butOK.TabStop = True
7         Else
8             butOK.ForeColor = &H80000011
9             txtbxDownArrow.ForeColor = &H80000011
10            butOK.TakeFocusOnClick = False
11            butOK.TabStop = False
12        End If

13        Exit Sub
ErrHandler:
14        Throw "#butOK_EnableOrDisable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TxtBxFilter_Exit(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler

2         butOK_EnableOrDisable

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TxtBxFilter_Exit (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TxtBxFilter_KeyDown
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : When hitting down arrow from the Filter Text box, make the first element
'             in the list box have focus.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TxtBxFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 40 Then        'Down key
3             If LstBxChoices.ListCount > 0 Then
4                 LstBxChoices.SetFocus
5                 LstBxChoices.ListIndex = 0        'select the first element of the list
6             End If
7         End If
8         If (KeyCode = 81 Or KeyCode = 88) And Shift = 4 Then        'Alt Q, Alt X because these are alternatives to double-clicking on a _
                                                                       spreadsheet - see method AltBacktickResponse
9             AssistConstructFilter
10        End If
11        CaptureKeys KeyCode, Shift

12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#TxtBxFilter_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
1         On Error GoTo ErrHandler

2         Cancel = True
3         butCancel_Click

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#UserForm_QueryClose (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub TxtBxExplanations_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         TextBoxMouseEventB TxtBxExplanations, Button, Shift, x, y, True, m_TheChoices, m_HelpBrowserMode, LstBxChoices, TxtBxFilter

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TxtBxExplanations_MouseDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InitialiseAsChooserDialog
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Populating and sizing elements on the form for when we use the form as a relatively simple pick-list - ShowSingleChoiceDialog
' -----------------------------------------------------------------------------------------------------------------------
Sub InitialiseAsChooserDialog(TheChoices As Variant, _
        TheExplanations, _
        TheExplanationTitles, _
        InitialFilter As String, _
        Title As String, _
        TopText As String, _
        Optional TheCategories As Variant, _
        Optional CategoryLabel As String, _
        Optional RegistryString, Optional SettingsMenu, _
        Optional InitialCategory As String)

1         On Error GoTo ErrHandler

          'Populate module-level variables
2         m_StandardSearchDescription = "&Search items only"
3         m_FullSearchDescription = "&Full Search in items and descriptions"

4         m_TheChoices = TheChoices
5         m_TheExplanations = TheExplanations
6         m_FilteredExplanations = TheExplanations
7         m_TheExplanationTitles = TheExplanationTitles
8         m_FilteredExplanationTitles = TheExplanationTitles
9         m_NumRecords = sNRows(m_TheChoices)
10        If sNCols(SettingsMenu) = 3 Then
11            m_SettingsMenu = SettingsMenu
12        End If

          'line below copes gracefully with case when m_TheExplanationTitles and/or m_TheExplanations are missing
13        m_TheFullTexts = sArrayConcatenate(m_TheChoices, vbLf, m_TheExplanationTitles, vbLf, m_TheExplanations)

14        m_ShowExplanations = Not (IsMissing(TheExplanations) Or IsEmpty(TheExplanations))

15        m_HelpBrowserMode = False
16        m_ShowCategories = Not IsMissing(TheCategories)
17        If RegistryString <> vbNullString Then
18            m_RegistrySection = "ShowSingleChoiceDialog-" & RegistryString
19        End If

20        If m_ShowCategories Then
21            m_TheCategories = TheCategories
22            m_TheCategoriesNoDupes = sArrayStack("All", sRemoveDuplicates(TheCategories, True))
23        End If

          'Lay out the controls
24        PositionControlsAsChooserDialog Me, TheChoices, TheExplanations, InitialFilter, InitialCategory, Title, TopText, CategoryLabel, m_ShowExplanations, _
              m_ShowCategories, m_TheCategoriesNoDupes, m_RegistrySection, m_SearchInMode, m_SearchHowMode

25        UpdateNumRecords Me, m_ShowCategories, m_NumRecords, m_SearchInMode, m_SearchHowMode, m_ShowExplanations, m_StandardSearchDescription, m_FullSearchDescription
26        UpdateExplanation Me, m_ShowExplanations, m_HelpBrowserMode, m_WebHelpAvailable, m_FilteredExplanations, m_FilteredExplanationTitles
27        butOK_EnableOrDisable

28        SetUpResizerChooserDialogMode Me, m_clsResizer, m_ShowExplanations, m_ShowCategories

29        Exit Sub
ErrHandler:
30        Throw "#InitialiseAsChooserDialog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThisAddinHelpAddress
' Author    : Philip Swannell
' Date      : 12-Nov-2015
' Purpose   : Stub function
' -----------------------------------------------------------------------------------------------------------------------
Function ThisAddinHelpAddress(FunctionName As String)
1         ThisAddinHelpAddress = "http://This_PageDoesNotExistYet"
End Function

