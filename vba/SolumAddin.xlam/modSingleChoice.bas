Attribute VB_Name = "modSingleChoice"
Option Explicit
Option Private Module
Const m_WebHelpAvailable = False        '<- PGS 21-Apr-2015. Switch this to True if we write web-based help for the functions such as exists for Martin's TigerLib
Const m_Max_Width_List As Long = 500
Const m_Max_Rows_List As Long = 40
Const m_Max_Rows_Explanations = 10
Const m_Min_Rows_List As Long = 10
Const v_gap As Long = 10
Const m_Left As Long = 12

'PGS 21-May-2019 moved code to here from frmSingleChoice since the code module of that form was very large, which may be bad for stability

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TextBoxMouseEventB
' Author    : Philip Swannell
' Date      : 04-May-2016
' Purpose   : Previous to 4-May-2016 this was common code with that used in MsgBoxPlus to
'             implement hyperlinks to web pages, but no longer common since here we also
'             implement "links" to other topics when using the dialog in help browser mode.
' -----------------------------------------------------------------------------------------------------------------------
Sub TextBoxMouseEventB(tb As Object, Button As Integer, Shift As Integer, x As Single, y As Single, IsDownClick As Boolean, ValidJumpTos As Variant, HelpBrowserMode As Boolean, LB As Object, FilterTextBox As Object)
          
          Dim Choices
          Dim CurrentWord As String
          Dim FaceIDs
          Dim Res
          Const chCopy = "&Copy         (Ctrl C)"        'Ctrl C will only work if we have also implemented a KeyDown event for the textbox
          Dim address As String
          Dim chGoTo As String
          Dim chJumpTo As String
          Dim MatchRes
          Dim NewSelStart As Long
1         On Error GoTo ErrHandler

2         Choices = CreateMissing()
3         FaceIDs = CreateMissing()

4         If Button = 2 And IsDownClick = True Then        'Right down-click
5             If tb.SelLength > 0 Then
6                 Choices = sArrayStack(Choices, chCopy)
7                 FaceIDs = sArrayStack(FaceIDs, 22)
8             End If
9         End If
10        If Button = 1 And IsDownClick = False Then        'Left up-click
11            GetCurrentWordOfTextBox CurrentWord, tb, NewSelStart
12            If LCase$(Left$(CurrentWord, 4)) = "www." Or _
                  LCase$(Left$(CurrentWord, 7)) = "http://" Or _
                  LCase$(Left$(CurrentWord, 8)) = "https://" Then
13                chGoTo = "&Go to " + CurrentWord
14                Choices = sArrayStack(Choices, chGoTo)
15                FaceIDs = sArrayStack(FaceIDs, 9026)
16                If LCase$(Left$(CurrentWord, 4)) = "www." Then
17                    address = "http://" + CurrentWord
18                Else
19                    address = CurrentWord
20                End If
21                tb.SelStart = NewSelStart
22                tb.SelLength = Len(CurrentWord)
23            ElseIf HelpBrowserMode Then
24                If IsNumeric(sMatch(CurrentWord, ValidJumpTos)) Then
25                    chJumpTo = "Help for " + CurrentWord
26                    Choices = sArrayStack(Choices, chJumpTo)
27                    FaceIDs = sArrayStack(FaceIDs, 49)
28                    tb.SelStart = NewSelStart
29                    tb.SelLength = Len(CurrentWord)
30                End If
31            End If
32        End If
33        If IsMissing(Choices) Then Exit Sub

          Dim PI As clsPositionInstructions

34        CreatePositionInstructions PI, tb, _
              (-tb.Width / 2 + x - 60) * fx, _
              (-tb.Height / 2 + y + 20) * fY

35        Res = ShowCommandBarPopup(Choices, FaceIDs, , , PI)

36        Select Case Res
              Case Unembellish(chCopy)
37                CopyStringToClipboard tb.SelText
38            Case Unembellish(chGoTo)
39                ThisWorkbook.FollowHyperlink address:=address, NewWindow:=True
40            Case Unembellish(chJumpTo)
41                MatchRes = sMatch(CurrentWord, sSubArray(LB.List, 1, 1))        'sSubArray necessary since sMatch does not cope with 0-based arrays
42                If IsNumeric(MatchRes) Then
43                    LB.ListIndex = MatchRes - 1
44                Else
45                    FilterTextBox.Value = CurrentWord
46                    MatchRes = sMatch(CurrentWord, sSubArray(LB.List, 1, 1))
47                    If IsNumeric(MatchRes) Then LB.ListIndex = MatchRes - 1
48                End If
49        End Select
50        Exit Sub
ErrHandler:
51        Throw "#TextBoxMouseEventB (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetUpResizerHelpBrowserMode
' Author    : Philip Swannell
' Date      : 28-Nov-2015
' Purpose   : Set up the Resizer class, including calculating the minimum size that we
'             will allow the form to shrink to and the relative speeds at which the
'             LstBxChoices and TxtBxExplanations should shrink as the user changes
'             the form's height. This version for use when the form is in use in Help Browser mode
' -----------------------------------------------------------------------------------------------------------------------
Sub SetUpResizerHelpBrowserMode(frm As frmSingleChoice, ByRef Resizer As clsFormResizer)

          Dim MaxShrinkageChoices As Double
          Dim MinWidth As Double
          Dim ShrinkageRatio As Double
          Dim WTagForChoices As String
          Dim WTagForExplanations As String

1         On Error GoTo ErrHandler

          'b) Calculate the maximum amount we will allow each of the two shrinkable controls to shrink
2         MaxShrinkageChoices = SafeMax(frm.LstBxChoices.Height - 53, 0)        '53 allow 5 lines of text to appear
          'c) Set the tags needed by the resizer
3         ShrinkageRatio = (frm.CmbBxCategories.Width - 10) / (frm.CmbBxCategories.Width - 10 + frm.TxtBxExplanations.Width)
          'a) Calculate the minimum width that we will allow the form to resize to:
4         If m_WebHelpAvailable Then
5             MinWidth = frm.butOK.Width + frm.butCancel.Width + m_Left * 3 + frm.Width - frm.InsideWidth + frm.lblHelp.Width
6         Else

              Dim ButtonGap

7             ButtonGap = frm.butCancel.Left - (frm.butOK.Left + frm.butOK.Width)
8             MinWidth = frm.Width - (ButtonGap - m_Left) / ShrinkageRatio
9         End If

10        WTagForChoices = CStr(ShrinkageRatio)
11        WTagForExplanations = CStr(1 - ShrinkageRatio)

12        frm.LstBxChoices.Tag = "H" + "W" + WTagForChoices
13        frm.TxtBxFilter.Tag = "W" + WTagForChoices
14        frm.CmbBxCategories.Tag = "W" + WTagForChoices
15        frm.lblSearchOptions.Tag = "L" + WTagForChoices
16        frm.butOK.Tag = "T"
17        frm.butCancel.Tag = "T" + "L" + WTagForChoices
18        frm.lblNumRecords.Tag = "T"
19        frm.lblExplanationTitle.Tag = "L" + WTagForChoices + "W" + WTagForExplanations
20        frm.TxtBxExplanations.Tag = "H" + "L" + WTagForChoices + "W" + WTagForExplanations

21        frm.txtbxDownArrow.Tag = "T"
22        frm.lblDivider.Tag = "T"
23        frm.lblHelp.Tag = "T"

          'd) Create the instance of the resizer class
24        CreateFormResizer Resizer

          'e) Tell it which form it's handling and set minimum height and width
25        Resizer.Initialise frm, frm.Height - MaxShrinkageChoices, MinWidth

26        Exit Sub
ErrHandler:
27        Throw "#SetUpResizerHelpBrowserMode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetUpResizerChooserDialogMode
' Author    : Philip Swannell
' Date      : 18-Nov-2013
' Purpose   : Set up the Resizer class, including calculating the minimum size that we
'             will allow the form to shrink to and the relative speeds at which the
'             LstBxChoices and TxtBxExplanations should shrink as the user changes
'             the form's height.
' -----------------------------------------------------------------------------------------------------------------------
Sub SetUpResizerChooserDialogMode(frm As frmSingleChoice, ByRef Resizer As clsFormResizer, ShowExplanations As Boolean, ShowCategories As Boolean)

          Dim HTagForChoices As String
          Dim HTagForExplanations As String
          Dim MaxShrinkageChoices As Double
          Dim MaxShrinkageExplanations As Double
          Dim MinWidth As Double
          Dim ShrinkageRatio As Double

1         On Error GoTo ErrHandler

          'a) Calculate the minimum width that we will allow the form to resize to:

2         MinWidth = frm.butCancel.Left + frm.butCancel.Width + frm.Width - frm.InsideWidth + 10

          'b) Calculate the maximum amount we will allow each of the two shrinkable controls to shrink
3         MaxShrinkageChoices = SafeMax(frm.LstBxChoices.Height - 53, 0)        '53 allow 5 lines of text to appear
4         If ShowExplanations Then
5             MaxShrinkageExplanations = SafeMax(frm.TxtBxExplanations.Height - 25, 0)        '25 allows two lines of text to appear
6         Else
7             MaxShrinkageExplanations = 0
8         End If

9         If (MaxShrinkageChoices = 0 And MaxShrinkageExplanations = 0) Or _
              Not ShowExplanations Then
10            HTagForChoices = vbNullString
11            HTagForExplanations = vbNullString
12        Else
              'Arrange that they both hit their minimum size (maximum shrinkage) at the same time.
13            ShrinkageRatio = MaxShrinkageChoices / (MaxShrinkageChoices + MaxShrinkageExplanations)
14            HTagForChoices = CStr(ShrinkageRatio)
15            HTagForExplanations = CStr(1 - ShrinkageRatio)
16        End If

          'c) Set the tags needed by the resizer
17        frm.LstBxChoices.Tag = "H" + HTagForChoices + "W"
18        frm.TxtBxFilter.Tag = "W"
19        frm.butOK.Tag = "T"
20        frm.butCancel.Tag = "T"
21        frm.lblNumRecords.Tag = "T"
22        frm.lblSearchOptions.Tag = "L"
23        If ShowExplanations Then
24            frm.lblExplanationTitle.Tag = "T" + HTagForChoices + "W"
25            frm.TxtBxExplanations.Tag = "T" + HTagForChoices + "WH" + HTagForExplanations
26            frm.txtbxDownArrow.Tag = "T"
27            frm.lblDivider.Tag = "T"
28        End If
29        If ShowCategories Then
30            frm.CmbBxCategories.Tag = "W"
31        End If

          'd) Create the instance of the resizer class
32        CreateFormResizer Resizer

          'e) Tell it which form it's handling and set minimum height and width
33        Resizer.Initialise frm, frm.Height - MaxShrinkageChoices - MaxShrinkageExplanations, MinWidth

34        Exit Sub
ErrHandler:
35        Throw "#SetUpResizerChooserDialogMode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PositionControlsAsChooserDialog
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Populating and sizing elements on the form for when we use the form as a relatively simple pick-list - ShowSingleChoiceDialog
' -----------------------------------------------------------------------------------------------------------------------
Sub PositionControlsAsChooserDialog(frm As frmSingleChoice, TheChoices As Variant, _
          TheExplanations, _
          InitialFilter As String, _
          InitialCategory As String, _
          Title As String, _
          TopText As String, _
          CategoryLabel As String, _
          ShowExplanations As Boolean, _
          ShowCategories As Boolean, _
          TheCategoriesNoDupes, _
          RegistrySection As String, _
          ByRef m_SearchInMode As Long, _
          ByRef m_SearchHowMode As Long)

          Dim Min_Width_Categories As Double
          Dim Min_Width_list As Double
          Dim NumRowsExplanations As Long
          Dim NumRowsToShow As Long

1         On Error GoTo ErrHandler

2         SetFonts frm, "Segoe UI", 9
3         frm.caption = Title

          'Horizontal postioning.
4         frm.LstBxChoices.List = TheChoices
5         frm.LstBxChoices.Width = frm.MinWidthForComboBox(frm.LstBxChoices, sArrayStack(TopText, TheChoices))
6         Min_Width_list = frm.butCancel.Left + frm.butCancel.Width - frm.butOK.Left
7         If ShowExplanations Then
8             frm.lblAutoSize.caption = sConcatenateStrings(TheExplanations, vbLf)
9             frm.lblAutoSize.Width = 1000
10            frm.lblAutoSize.AutoSize = False
11            frm.lblAutoSize.AutoSize = True
12            Min_Width_list = SafeMax(Min_Width_list, SafeMin(310, frm.lblAutoSize.Width))
13            frm.lblAutoSize.caption = vbNullString
14        End If

15        frm.LstBxChoices.Left = m_Left
16        If frm.LstBxChoices.Width > m_Max_Width_List Then
17            frm.LstBxChoices.Width = m_Max_Width_List
18        ElseIf frm.LstBxChoices.Width < Min_Width_list Then
19            frm.LstBxChoices.Width = Min_Width_list
20        End If

21        If TopText <> vbNullString Then
22            frm.lblTopText.Width = frm.LstBxChoices.Width
23            frm.lblTopText.caption = TopText
24            frm.lblTopText.AutoSize = False
25            frm.lblTopText.AutoSize = True
26            frm.lblTopText.Left = m_Left
27        Else
28            frm.lblTopText.Visible = False
29        End If

30        frm.TxtBxFilter.Height = 20
31        If ShowExplanations Then
32            frm.TxtBxExplanations.Left = m_Left
33            frm.TxtBxExplanations.Width = frm.LstBxChoices.Width
34            frm.TxtBxExplanations.Visible = True
35            frm.lblExplanationTitle.Visible = True
36        Else
37            frm.lblExplanationTitle.Visible = False
38            frm.TxtBxExplanations.Visible = False
39        End If

40        If ShowCategories Then
41            With frm.lblCategory
42                .caption = CategoryLabel
43                .Width = 200
44                .AutoSize = False
45                .AutoSize = True
46                .Left = m_Left + 8.5
47            End With
48            With frm.CmbBxCategories
49                .List = TheCategoriesNoDupes
50                .Left = frm.lblCategory.Left + frm.lblCategory.Width + 3
51                .Height = frm.TxtBxFilter.Height
52                .Width = frm.LstBxChoices.Left + frm.LstBxChoices.Width - .Left
53                .Value = "All"
54                Min_Width_Categories = frm.MinWidthForComboBox(frm.CmbBxCategories)
55                If .Width < Min_Width_Categories Then
                      Dim nudge As Double
56                    nudge = Min_Width_Categories - .Width
57                    .Width = Min_Width_Categories
58                    frm.LstBxChoices.Width = frm.LstBxChoices.Width + nudge
59                End If

60            End With
61        Else
62            frm.CmbBxCategories.Visible = False
63            frm.lblCategory.Visible = False
64        End If

65        frm.TxtBxFilter.Left = m_Left
66        frm.TxtBxFilter.Width = frm.LstBxChoices.Width
67        frm.Width = frm.LstBxChoices.Left + frm.LstBxChoices.Width + m_Left + 4
68        frm.lblNumRecords.Left = m_Left

          'Vertical positioning
69        If frm.lblTopText.Visible Then
70            frm.lblTopText.Top = v_gap
71            frm.TxtBxFilter.Top = frm.lblTopText.Top + frm.lblTopText.Height + v_gap - 7
72        Else
73            frm.TxtBxFilter.Top = v_gap
74        End If

75        If ShowCategories Then
76            frm.CmbBxCategories.Top = frm.TxtBxFilter.Top + frm.TxtBxFilter.Height + v_gap
77            frm.LstBxChoices.Top = frm.CmbBxCategories.Top + frm.CmbBxCategories.Height + v_gap
78            frm.lblCategory.Top = frm.CmbBxCategories.Top + frm.CmbBxCategories.Height / 2 - frm.lblCategory.Height / 2 - 1
79        Else
80            frm.LstBxChoices.Top = frm.TxtBxFilter.Top + frm.TxtBxFilter.Height + v_gap
81        End If

82        NumRowsToShow = sNRows(TheChoices)
83        If ShowExplanations Then
84            frm.TxtBxExplanations.Height = GoodHeight(frm.TxtBxExplanations, frm.TxtBxExplanations.Width, TheExplanations)
85            If frm.TxtBxExplanations.Height > m_Max_Rows_Explanations * 10 Then
86                frm.TxtBxExplanations.Height = m_Max_Rows_Explanations * 10
87            End If
88            NumRowsExplanations = frm.TxtBxExplanations.Height / 10
89        End If

90        If NumRowsToShow < m_Min_Rows_List Then
91            NumRowsToShow = m_Min_Rows_List
92        ElseIf NumRowsToShow > m_Max_Rows_List - NumRowsExplanations - 1 Then
93            NumRowsToShow = m_Max_Rows_List - NumRowsExplanations - 1
94        End If
          Dim Multiplier As Double
          Dim Offset As Double
95        VerticalSizing frm, Multiplier, Offset

96        frm.LstBxChoices.Height = NumRowsToShow * Multiplier + Offset
97        frm.LstBxChoices.IntegralHeight = False
98        frm.LstBxChoices.IntegralHeight = True

99        If ShowExplanations Then
100           frm.TxtBxExplanations.Top = frm.LstBxChoices.Top + frm.LstBxChoices.Height + v_gap + 7
101           frm.lblExplanationTitle.Left = m_Left
102           frm.lblExplanationTitle.Top = frm.TxtBxExplanations.Top - 14
103           frm.lblExplanationTitle.Height = 12
104           frm.lblExplanationTitle.Width = frm.TxtBxExplanations.Width
105           frm.butOK.Top = frm.TxtBxExplanations.Top + frm.TxtBxExplanations.Height + v_gap
106       Else
107           frm.butOK.Top = frm.LstBxChoices.Top + frm.LstBxChoices.Height + v_gap
108       End If
109       frm.butOK.Left = m_Left
110       frm.butCancel.Top = frm.butOK.Top
111       frm.butCancel.Height = frm.butOK.Height
112       frm.butCancel.Width = frm.butOK.Width
113       frm.butCancel.Left = frm.butOK.Left + frm.butOK.Width + 10

114       frm.lblNumRecords.Top = frm.butCancel.Top + frm.butCancel.Height + v_gap
115       frm.lblNumRecords.Height = 10
116       frm.Height = frm.lblNumRecords.Top + frm.lblNumRecords.Height + v_gap + (frm.Height - frm.InsideHeight + 6.35)

          'Position the Search Options label.
117       With frm.lblSearchOptions
118           .Top = frm.TxtBxFilter.Top
119           .Height = frm.TxtBxFilter.Height
120           .Width = .Height
121           .Left = m_Left + frm.TxtBxFilter.Width - .Width
122           frm.TxtBxFilter.Width = frm.TxtBxFilter.Width - .Width + 0.5
123       End With

124       frm.lblHelp.Visible = False

125       If RegistrySection <> vbNullString Then
              'Code in this section must be in synch with corresponding calls to SaveSettings in method lblSearchOptions_Click
126           m_SearchInMode = IIf(GetSetting(gAddinName, RegistrySection, "SearchInMode", "ItemsOnly") = "ItemsOnly", 0, 1)
127           Select Case GetSetting(gAddinName, RegistrySection, "SearchHowMode", "Simple")
                  Case "Simple"
128                   m_SearchHowMode = 0
129               Case "StartsWith"
130                   m_SearchHowMode = 1
131               Case "Pattern"
132                   m_SearchHowMode = 2        'No longer support pattern matching, flip to RegExp
133               Case "RegExp"
134                   m_SearchHowMode = 2
135           End Select
136       Else
137           m_SearchInMode = 0
138           m_SearchHowMode = 0
139       End If

140       If Not ShowExplanations Then m_SearchInMode = 0
          'Have to set the value of TxtBxFilter after we have set m_SearchInMode and m_SearchInMode
141       frm.TxtBxFilter.Value = InitialFilter
142       If InitialCategory <> "" Then
143           frm.CmbBxCategories.Value = InitialCategory
144       End If
145       frm.lblDivider.Visible = False
146       frm.txtbxDownArrow.Visible = False

147       Exit Sub
ErrHandler:
148       Throw "#PositionControlsAsChooserDialog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub PositionControlsAsHelpBrowser(frm As frmSingleChoice, TheChoices As Variant, _
        TheExplanations, _
        TheExplanationTitles, _
        Title As String, _
        TopText As String, _
        TheCategories As Variant, _
        CategoryLabel As String, _
        RegistryString As String, _
        RegistrySection As String, _
        TheCategoriesNoDupes, _
        ByRef m_SearchInMode As Long, _
        ByRef m_SearchHowMode As Long)

          Dim NumRowsToShow As Long

1         On Error GoTo ErrHandler

2         SetFonts frm, "Segoe UI", 9
3         frm.caption = Title

4         frm.LstBxChoices.List = TheChoices
5         frm.TxtBxFilter.Value = vbNullString

          'Horizontal postioning. Use frm.lblAutoSize's autosize method to set the _
           width of frm.LstBxChoices, which does not have its own AutoSize method.
6         frm.lblAutoSize.caption = sConcatenateStrings(TheChoices, vbLf) + vbLf + sConcatenateStrings(TheCategoriesNoDupes, vbLf)
7         frm.lblAutoSize.Width = m_Max_Width_List + 100
8         frm.lblAutoSize.AutoSize = True

9         frm.LstBxChoices.Width = frm.lblAutoSize.Width + 60
10        frm.lblAutoSize.caption = vbNullString
11        frm.lblAutoSize.Visible = False

12        frm.LstBxChoices.Left = m_Left

13        frm.lblTopText.Width = frm.LstBxChoices.Width
14        frm.lblTopText.caption = TopText
15        frm.lblTopText.AutoSize = False
16        frm.lblTopText.AutoSize = True
17        frm.lblTopText.Left = m_Left

18        frm.TxtBxExplanations.Left = m_Left
19        frm.TxtBxExplanations.Width = frm.LstBxChoices.Width
20        frm.TxtBxExplanations.Visible = True
21        frm.lblExplanationTitle.Visible = True

22        With frm.lblCategory
23            .caption = CategoryLabel
24            .Width = 200
25            .AutoSize = False
26            .AutoSize = True
27            .Left = m_Left
28        End With
29        With frm.CmbBxCategories
30            .Left = frm.lblCategory.Left + frm.lblCategory.Width + 3
31            .Height = frm.TxtBxFilter.Height
32            .Width = frm.LstBxChoices.Left + frm.LstBxChoices.Width - .Left
33            .List = TheCategoriesNoDupes
34            .Value = "All"
35        End With

36        frm.TxtBxFilter.Left = m_Left
37        frm.TxtBxFilter.Width = frm.LstBxChoices.Width
38        frm.lblNumRecords.Left = m_Left

          'Vertical positioning

39        frm.lblTopText.Top = v_gap
40        frm.TxtBxFilter.Top = frm.lblTopText.Top + frm.lblTopText.Height + v_gap - 7
41        frm.CmbBxCategories.Top = frm.TxtBxFilter.Top + frm.TxtBxFilter.Height + v_gap
42        frm.LstBxChoices.Top = frm.CmbBxCategories.Top + frm.CmbBxCategories.Height + v_gap
43        frm.lblCategory.Top = frm.CmbBxCategories.Top + frm.CmbBxCategories.Height / 2 - frm.lblCategory.Height / 2 - 1

          'Position the Search Options label.
44        With frm.lblSearchOptions
45            .Top = frm.TxtBxFilter.Top
46            .Height = frm.TxtBxFilter.Height
47            .Width = .Height
48            .Left = m_Left + frm.TxtBxFilter.Width - .Width
49        End With

50        NumRowsToShow = sNRows(TheChoices)

51        If NumRowsToShow < m_Min_Rows_List Then
52            NumRowsToShow = m_Min_Rows_List
53        ElseIf NumRowsToShow > m_Max_Rows_List Then
54            NumRowsToShow = m_Max_Rows_List
55        End If
          Dim Multiplier As Double
          Dim Offset As Double
56        VerticalSizing frm, Multiplier, Offset

57        frm.LstBxChoices.Height = NumRowsToShow * Multiplier + Offset
58        frm.LstBxChoices.IntegralHeight = False
59        frm.LstBxChoices.IntegralHeight = True

60        With frm.TxtBxExplanations
61            .BackColor = frm.LstBxChoices.BackColor
62            .Top = frm.TxtBxFilter.Top
63            .Left = frm.LstBxChoices.Left + frm.LstBxChoices.Width + 20
64            .Height = frm.LstBxChoices.Top + frm.LstBxChoices.Height - .Top
65            .Width = 400
66        End With

67        With frm.lblExplanationTitle
68            .Left = frm.TxtBxExplanations.Left
69            .Top = frm.lblTopText.Top

70            .Height = frm.lblTopText.Height
71            .Width = frm.TxtBxExplanations.Width
72        End With

73        frm.butOK.Top = frm.LstBxChoices.Top + frm.LstBxChoices.Height + v_gap
74        frm.butOK.caption = "Insert Function"        '"Enter formula" + vbLf + "on sheet"
75        frm.butOK.Width = 100
76        frm.butOK.Left = m_Left
77        frm.butCancel.Top = frm.butOK.Top
78        frm.butCancel.Height = frm.butOK.Height
79        frm.butCancel.Width = frm.butOK.Width
80        frm.butCancel.Left = frm.TxtBxExplanations.Left

81        frm.lblNumRecords.Top = frm.butCancel.Top + frm.butCancel.Height + v_gap
82        frm.lblNumRecords.Height = 10

83        frm.Height = frm.lblNumRecords.Top + frm.lblNumRecords.Height + v_gap + (frm.Height - frm.InsideHeight + 6.35)
84        frm.Width = frm.TxtBxExplanations.Left + frm.TxtBxExplanations.Width + m_Left + 10

85        With frm.txtbxDownArrow
86            .BorderStyle = fmBorderStyleNone
87            .Top = frm.butOK.Top + 1.5
88            .Width = frm.butOK.Width / 6
89            .Height = frm.butOK.Height - 4
90            .Left = frm.butOK.Left + frm.butOK.Width - .Width - 3
              'And position the dividing line between the two parts of the button
91            frm.lblDivider.Top = frm.butOK.Top + 4
92            frm.lblDivider.Left = .Left
93            frm.lblDivider.Height = frm.butOK.Height - 8
94            frm.lblDivider.BorderStyle = fmBorderStyleNone
95        End With

96        If m_WebHelpAvailable Then
97            With frm.lblHelp
98                .Top = frm.butCancel.Top + frm.butCancel.Height / 2 - .Height / 2
99                .Left = frm.butCancel.Left + frm.butCancel.Width + 10
100               .Visible = True
101           End With
102       Else
103           frm.lblHelp.Visible = False
104       End If

105       If RegistrySection <> vbNullString Then
              'Code in this section must be in synch with corresponding calls to SaveSettings in method frm.lblSearchOptions_Click
106           m_SearchInMode = IIf(GetSetting(gAddinName, RegistrySection, "SearchInMode", "ItemsOnly") = "ItemsOnly", 0, 1)
107           Select Case GetSetting(gAddinName, RegistrySection, "SearchHowMode", "Simple")
                  Case "Simple"
108                   m_SearchHowMode = 0
109               Case "StartsWith"
110                   m_SearchHowMode = 1
111               Case "Pattern"
112                   m_SearchHowMode = 2        ' no longer support Pattern, flip to RegExp
113               Case "RegExp"
114                   m_SearchHowMode = 2
115           End Select
116       Else
117           m_SearchInMode = 0
118           m_SearchHowMode = 0
119       End If

120       Exit Sub
ErrHandler:
121       Throw "#PositionControlsAsHelpBrowser (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetFonts
' Author    : Philip Swannell
' Date      : 06-Jun-2016
' Purpose   : Set the font of all controls
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SetFonts(frm As frmSingleChoice, FontName As String, FontSize As Double)
1         On Error GoTo ErrHandler
          Dim c As control
2         For Each c In frm.Controls
3             If Not c Is frm.txtbxDownArrow Then
4                 c.Font.Name = FontName
5                 If c Is frm.LstBxChoices Then
6                     c.Font.Size = FontSize - 0.5
7                 Else
8                     c.Font.Size = FontSize
9                 End If
10            End If
11        Next
12        Exit Sub
ErrHandler:
13        Throw "#SetFonts (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : VerticalSizing
' Author    : Philip Swannell
' Date      : 07-Jun-2016
' Purpose   : LstBoxChoices has no autosize method, we use this method to determine the
'             required height as a linear function of the number of rows displayed by using
'             LblAutoSize, which does have the AutoSize property
' -----------------------------------------------------------------------------------------------------------------------
Private Sub VerticalSizing(frm As frmSingleChoice, ByRef Multiplier As Double, Offset As Double)
1         On Error GoTo ErrHandler
          Dim x1 As Double
          Dim X2 As Double
          Dim y1 As Double
          Dim y2 As Double

2         With frm.lblAutoSize
3             .caption = vbNullString
4             .caption = "XYZ"
5             .AutoSize = False
6             .Width = 100
7             .AutoSize = True
8             x1 = 1
9             y1 = .Height
10            .caption = "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ" + vbCr + "XYZ"
11            .AutoSize = False
12            .Width = 100
13            .AutoSize = True
14            X2 = 10
15            y2 = .Height
16        End With
17        Multiplier = (y2 - y1) / (X2 - x1)
18        Offset = y1 - Multiplier * x1 + 4

19        Exit Sub
ErrHandler:
20        Throw "#VerticalSizing (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GoodHeight
' Author    : Philip Swannell
' Date      : 10-Nov-2013
' Purpose   : Determine the height of a textbox of a given width so that it can display any
'             of the strings contained in PossibleCaptions without need for ScrollBars
' -----------------------------------------------------------------------------------------------------------------------
Private Function GoodHeight(TheTextBox As Object, LabelWidth As Double, PossibleCaptions As Variant)

          Dim i As Long
          Dim newHeight As Double
          Dim origAutoSize As Boolean
          Dim OrigHeight As Double
          Dim OrigWidth As Double

1         On Error GoTo ErrHandler

2         OrigHeight = TheTextBox.Height
3         OrigWidth = TheTextBox.Width
4         origAutoSize = TheTextBox.AutoSize

5         Force2DArray PossibleCaptions
6         With TheTextBox
7             For i = 1 To sNRows(PossibleCaptions)
8                 .text = PossibleCaptions(i, 1)
9                 .Width = LabelWidth
10                .AutoSize = False
11                .AutoSize = True
12                If .Height > newHeight Then newHeight = .Height
13            Next i
14        End With
15        TheTextBox.AutoSize = origAutoSize
16        TheTextBox.Height = OrigHeight
17        TheTextBox.Width = OrigWidth

18        GoodHeight = newHeight
19        Exit Function
ErrHandler:
20        Throw "#GoodHeight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UpdateNumRecords
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : When some but not all records are filtered then show a message about how many
'             are filtered, and give hints about the search method in use and what the user
'             might do to if there are no hits.
' -----------------------------------------------------------------------------------------------------------------------
Sub UpdateNumRecords(frm As frmSingleChoice, ShowCategories As Boolean, NumRecords As Long, SearchInMode As Long, SearchHowMode As Long, ShowExplanations As Boolean, StandardSearchDescription As String, FullSearchDescription As String)
          Dim FilterByCategory As Boolean
          Dim Message As String
          Const Hint1 = vbLf + "No hits? See Search Options (F2)."
          Const Hint2 = vbLf + "No hits? That regular expression has a syntax error!"

          Dim NoHits As Boolean

1         On Error GoTo ErrHandler

2         NoHits = (frm.LstBxChoices.ListCount = 0)

3         If ShowCategories Then
4             If frm.CmbBxCategories.Value <> "All" Then
5                 FilterByCategory = True
6             End If
7         End If

8         If Len(frm.TxtBxFilter.text) = 0 And Not FilterByCategory Then
9             Message = vbNullString
10        Else
11            Message = Format$(frm.LstBxChoices.ListCount, "###,##0") + "/" & Format$(NumRecords, "###,###") & "."
12            If SearchInMode = 0 Then
13                If ShowExplanations Then
14                    Message = Message + " " + Unembellish(StandardSearchDescription) + ", "
15                End If
16            Else
17                Message = Message + " " + Unembellish(FullSearchDescription) + ", "
18            End If

19            Select Case SearchInMode & SearchHowMode
                  Case "00"        'Standard Search, Simple Match
20                    Message = Message + " Simple match."
21                    If NoHits Then Message = Message + Hint1
22                Case "01"        'Standard Search, StartsWith Match
23                    Message = Message + " 'Starts with' match."
24                    If NoHits Then Message = Message + Hint1
25                Case "02"
26                    Message = Message + " Regular Expression match."
27                    If NoHits Then
28                        If VarType(sIsRegMatch(frm.TxtBxFilter.Value, "Foo", False)) <> vbBoolean Then
29                            Message = Message + Hint2
30                        Else
31                            Message = Message + Hint1
32                        End If
33                    End If
34                Case "10"        'Full Search, Simple Match
35                    Message = Message + " Simple match."
36                Case "11"        'Full Search, StartsWith
37                    Message = Message + " 'Starts with' match."
38                    If NoHits Then Message = Message + Hint1
39                Case "12"        'Full Search Regular Expression
40                    Message = Message + " Regular Expression match."
41                    If NoHits Then
42                        If VarType(sIsRegMatch(frm.TxtBxFilter.Value, "Foo", False)) <> vbBoolean Then
43                            Message = Message + Hint2
44                        Else
45                            Message = Message + Hint1
46                        End If
47                    End If
48            End Select
49        End If

50        If frm.lblNumRecords.caption <> Message Then
51            frm.lblNumRecords.caption = Message
52            frm.lblNumRecords.Width = 500
53            frm.lblNumRecords.AutoSize = False
54            frm.lblNumRecords.AutoSize = True
55        End If

56        Exit Sub
ErrHandler:
57        Throw "#UpdateNumRecords (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UpdateExplanation
' Author    : Philip Swannell
' Date      : 10-Nov-2013
' Purpose   : Populate the text box showing the explanatory text for the user's current choice.
' -----------------------------------------------------------------------------------------------------------------------
Sub UpdateExplanation(frm As frmSingleChoice, ShowExplanations As Boolean, HelpBrowserMode As Boolean, WebHelpAvailable As Boolean, FilteredExplanations As Variant, FilteredExplanationTitles)
1         On Error GoTo ErrHandler

2         If ShowExplanations Then

3             If Not IsNull(frm.LstBxChoices.Value) Then
4                 frm.TxtBxExplanations.text = FilteredExplanations(frm.LstBxChoices.ListIndex + 1, 1)
5                 frm.lblExplanationTitle.caption = FilteredExplanationTitles(frm.LstBxChoices.ListIndex + 1, 1)
6                 frm.TxtBxExplanations.Visible = True
7                 KickScrollBars frm, ShowExplanations
8             ElseIf frm.LstBxChoices.ListCount = 1 Then
9                 frm.TxtBxExplanations.text = FilteredExplanations(1, 1)
10                frm.lblExplanationTitle.caption = FilteredExplanationTitles(1, 1)
11                frm.TxtBxExplanations.Visible = True
12                KickScrollBars frm, ShowExplanations
13            Else
14                frm.TxtBxExplanations.text = vbNullString
15                If HelpBrowserMode Then
16                    frm.lblExplanationTitle.caption = "Select a topic to see help"
17                    frm.TxtBxExplanations.Visible = True
18                Else
19                    frm.lblExplanationTitle.caption = vbNullString
20                    frm.TxtBxExplanations.Visible = False
21                End If
22            End If
23        End If

24        If HelpBrowserMode Then
25            If m_WebHelpAvailable Then
26                If Not IsNull(frm.LstBxChoices.Value) Then
27                    frm.lblHelp.Enabled = True
28                    frm.lblHelp.ControlTipText = "Click or F1 to go to " & frm.ThisAddinHelpAddress(frm.LstBxChoices.Value)
29                ElseIf frm.LstBxChoices.ListCount = 1 Then
30                    frm.lblHelp.Enabled = True
31                    frm.lblHelp.ControlTipText = frm.ThisAddinHelpAddress(frm.LstBxChoices.List(0))
32                Else
33                    frm.lblHelp.Enabled = False
34                    frm.lblHelp.ControlTipText = vbNullString
35                End If
36            End If
37        End If

38        Exit Sub
ErrHandler:
39        Throw "#UpdateExplanation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' Procedure : KickScrollBars
' Author    : Philip Swannell
' Date      : 10-Nov-2013
' Purpose   : The ScrollBars on the TxtBxExplanations can be reluctant to appear\disappear
'             this method makes them do so.
' -----------------------------------------------------------------------------------------------------------------------
Sub KickScrollBars(frm As frmSingleChoice, ShowExplanations As Boolean)
1         On Error GoTo ErrHandler
          Dim oldActive As control
2         If ShowExplanations Then
3             If frm.TxtBxExplanations.Visible Then
4                 If frm.ActiveControl Is Nothing Then Exit Sub
5                 Set oldActive = frm.ActiveControl
6                 frm.TxtBxExplanations.SetFocus
7                 frm.TxtBxExplanations.SelStart = 0
8                 frm.TxtBxExplanations.SelLength = 0
9                 oldActive.SetFocus
10            End If
11        End If
12        Exit Sub
ErrHandler:
13        Throw "#KickScrollBars (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
