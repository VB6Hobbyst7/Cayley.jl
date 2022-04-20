VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSelectTrades 
   Caption         =   "Select Trades from File"
   ClientHeight    =   4080
   ClientLeft      =   98
   ClientTop       =   434
   ClientWidth     =   5257
   OleObjectBlob   =   "FrmSelectTrades.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSelectTrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' -----------------------------------------------------------------------------------------------------------------------
' Module    : FrmSelectTrades
' Author    : Philip Swannell
' Date      : 20-Dec-2016
' Purpose   : A dialog for use from SCRiPT.xlsm to allow the user to choose parameters
'             for trade filtration that are exactly analagous to the filtration used
'             from the Cayley workbook.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_twb As Workbook
Private m_LinesBook As Workbook
Public ButtonClicked As String
Dim m_clsResizer As clsFormResizer
Dim m_ResizerInitialised As Boolean

Sub Initialise(twb As Workbook, FilterBy1 As String, Filter1Value As Variant, FilterBy2 As String, Filter2Value As Variant, _
               IncludeAssetClasses As String, CurrenciesToInclude, CompressTrades As Boolean, LinesBook As Workbook)

          Dim AllowedFilterBy As Variant

1         On Error GoTo ErrHandler
2         Set m_twb = twb
3         Set m_LinesBook = LinesBook
4         TextBoxFile.Value = twb.Name
5         AllowedFilterBy = GetColumnFromTradesWorkbook("AllHeaders", False, 0, True, True, twb, twb.Worksheets(1), Date)
6         AllowedFilterBy = sArrayStack("None", AllowedFilterBy)
7         ComboFilterBy1.List = AllowedFilterBy
8         ComboFilterBy1 = FilterBy1
9         ComboFilterBy2.List = AllowedFilterBy
10        ComboFilterBy2.Value = FilterBy2
11        TextBoxFilter1Value.Value = Filter1Value
12        TextBoxFilter2Value.Value = Filter2Value
13        ComboAssetClasses.List = sArrayStack("Rates and Fx", "Fx", "Rates")
14        ComboAssetClasses.Value = IncludeAssetClasses
15        TextBoxCurrencies.Value = CurrenciesToInclude
16        CheckBoxCompress.Value = CompressTrades

          'Set up resizer
17        TextBoxFile.Tag = "W"
18        ComboFilterBy1.Tag = "W"
19        TextBoxFilter1Value.Tag = "W"
20        ComboFilterBy2.Tag = "W"
21        TextBoxFilter2Value.Tag = "W"
22        ComboAssetClasses.Tag = "W"
23        TextBoxCurrencies.Tag = "W"
24        CheckBoxCompress.Tag = "W"
25        butDots1.Tag = "L"
26        butDots2.Tag = "L"
27        butDots3.Tag = "L"
28        CreateFormResizer m_clsResizer
29        m_clsResizer.Initialise Me, Me.Height, Me.Width, Me.BackColor
30        m_ResizerInitialised = True
31        RefreshControls
32        Exit Sub
ErrHandler:
33        SomethingWentWrong "#Initialise (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FilterValueDblClick
' Author    : Philip Swannell
' Date      : 20-Dec-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function FilterValueDblClick(FilterBy As String, twb As Workbook, ControlToSet As Control, LinesBook As Workbook)
          Dim EnableFlags As Variant
          Dim FaceIDs As Variant
          Dim Filters As Variant
          Dim RegKey As String
          Dim TheChoices As Variant
          Const chAdvanced = "Advanced..."
          Dim chPickMany As String
          Dim chPickOne As String
          Const MAXMRUsToShow = 9
          Dim Alternatives
          Dim Annotate As Boolean
          Dim ExistingValue As String
          Dim InitialChoices As Variant
          Dim NewValue As Variant
          Dim Res
          Dim UseRegEx As Boolean

1         Annotate = UCase(CStr(FilterBy)) = "CPTY_PARENT"
2         chPickOne = "Pick one " + FilterBy + "...      "
3         chPickMany = "Pick &multiple " + FilterBy + "s ..."

4         On Error GoTo ErrHandler
5         ExistingValue = ControlToSet.Value

6         If (LCase(FilterBy)) = "none" Then
7             NewValue = "None"
8             GoTo EarlyExit
9         Else
10            RegKey = "CayleyFilterBy" & FilterBy        'Shared registry key with double-clicking in Cayley workbook PFE sheet
11            GetMRUFilters RegKey, Filters, TheChoices, FaceIDs, EnableFlags
12            If Not IsMissing(Filters) Then
13                If sNRows(Filters) > MAXMRUsToShow Then
14                    TheChoices = sSubArray(TheChoices, 1, 1, MAXMRUsToShow)

15                    FaceIDs = sSubArray(FaceIDs, 1, 1, MAXMRUsToShow)
16                    EnableFlags = sSubArray(EnableFlags, 1, 1, MAXMRUsToShow)
17                End If
18            End If
19            TheChoices = sArrayStack(TheChoices, "--&" & chPickOne, chPickMany, "--&" & chAdvanced)
20            If Annotate Then TheChoices = AnnotateBankNames(TheChoices, True, LinesBook, True)
21            FaceIDs = sArrayStack(FaceIDs, 447, 448, 502)
22            EnableFlags = sArrayStack(EnableFlags, True, True, True)
              Dim PopUpRes
23            PopUpRes = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , , False)
24            Select Case PopUpRes
              Case "#Cancel!"
25                NewValue = Empty
26                GoTo EarlyExit
27            Case chPickOne, chAdvanced, Unembellish(chPickMany)
28                Alternatives = GetColumnFromTradesWorkbook(FilterBy, False, 0, True, True, twb, twb.Worksheets(1), Date)
29                Alternatives = sRemoveDuplicates(Alternatives, True)
30                Alternatives = sArrayMakeText(Alternatives)
31                UseRegEx = (PopUpRes = chAdvanced)
32            Case Else
33                If Annotate Then PopUpRes = AnnotateBankNames(PopUpRes, False, LinesBook, True)
34                NewValue = CStr(PopUpRes)
35                GoTo EarlyExit
36            End Select
37            Select Case UCase(FilterBy)
              Case "CPTY_PARENT"
38                If PopUpRes = chPickOne Then
39                    Alternatives = AnnotateBankNames(Alternatives, True, LinesBook)
40                    Res = ShowSingleChoiceDialog(Alternatives, , , , , "Select Parent Counterparty", , Me)
41                    If IsEmpty(Res) Then
42                        GoTo EarlyExit
43                    Else
44                        NewValue = AnnotateBankNames(Res, False, LinesBook)
45                        GoTo EarlyExit
46                    End If
47                ElseIf PopUpRes = Unembellish(chPickMany) Then
                      Dim MiddleCaption As String
48                    MiddleCaption = "All with lines"
49                    If Len(CStr(ControlToSet.Value)) > 0 Then
50                        InitialChoices = sTokeniseString(CStr(ControlToSet.Value), "|")
51                        InitialChoices = sArrayLeft(InitialChoices, -1)
52                        InitialChoices = sArrayRight(InitialChoices, -1)
53                        InitialChoices = sRegExpFromLiteral(InitialChoices, True)
54                        InitialChoices = AnnotateBankNames(InitialChoices, True, LinesBook)
55                    End If
TryAgain:
56                    Res = ShowMultipleChoiceDialog(AnnotateBankNames(Alternatives, True, LinesBook), InitialChoices, "Choose Banks", , , , "OK", , Array(False, True), MiddleCaption, ButtonClicked)
57                    If sIsErrorString(Res) Then
58                        GoTo EarlyExit
59                    ElseIf ButtonClicked = MiddleCaption Then
60                        InitialChoices = sCompareTwoArrays(Alternatives, GetColumnFromLinesBook("CPTY_PARENT", LinesBook), "Common")
61                        If sNRows(InitialChoices) > 1 Then
62                            InitialChoices = AnnotateBankNames(sSubArray(InitialChoices, 2), True, LinesBook)
63                        Else
64                            InitialChoices = CreateMissing()
65                        End If

66                        GoTo TryAgain
67                    End If
68                    If sArraysIdentical(Res, "#User Cancel!") Then
69                        GoTo EarlyExit
70                    Else
71                        NewValue = sConcatenateStrings(sArrayConcatenate("^", sRegExpFromLiteral(AnnotateBankNames(Res, False, LinesBook)), "$"), "|")
72                        GoTo EarlyExit
73                    End If
74                End If
75            End Select
76        End If

77        If UseRegEx Then
78            Res = ShowRegularExpressionDialog(IIf(CStr(ExistingValue) = "None", "", ExistingValue), _
                                                FilterBy, Alternatives, Me, "Filter Trades", _
                                                "Include trades for which:", False, RegKey, True)
79            If Res = "#User Cancel!" Then Res = Empty
80            NewValue = Res
81        ElseIf PopUpRes = Unembellish(chPickMany) Then
82            If Len(CStr(ControlToSet.Value)) > 0 Then
83                InitialChoices = sTokeniseString(CStr(ControlToSet.Value), "|")
84                InitialChoices = sArrayLeft(InitialChoices, -1)
85                InitialChoices = sArrayRight(InitialChoices, -1)
86                InitialChoices = sRegExpFromLiteral(InitialChoices, True)
87            End If
88            Res = ShowMultipleChoiceDialog(Alternatives, InitialChoices, "Select " + FilterBy + "s", , , Me, , , False)
89            If sArraysIdentical(Res, "#User Cancel!") Then
90                NewValue = Empty
91            Else
92                NewValue = sConcatenateStrings(sArrayConcatenate("^", sRegExpFromLiteral(Res), "$"), "|")
93            End If
94        ElseIf sNRows(Alternatives) < 20 Then
95            Res = ShowCommandBarPopup(Alternatives, , , ExistingValue, Me)
96            If Res = "#Cancel!" Then Res = Empty
97            NewValue = Res
98        Else
99            Alternatives = sArrayMakeText(Alternatives)
100           Res = ShowSingleChoiceDialog(Alternatives, , , , , "Select " + FilterBy, , Me)
101           NewValue = Res
102       End If

EarlyExit:
103       If Not IsEmpty(NewValue) Then
104           ControlToSet.Value = NewValue
105           RefreshControls
106       End If

107       AddCayleyFiltersToMRU twb, ComboFilterBy1.Value, CastNumberStringToNumber(TextBoxFilter1Value.Value), _
                                ComboFilterBy2.Value, CastNumberStringToNumber(TextBoxFilter2Value.Value), Date
108       Exit Function
ErrHandler:
109       Throw "#FilterValueDblClick (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Sub butCancel_Click()
1         On Error GoTo ErrHandler
2         ButtonClicked = "Cancel"
3         HideForm Me
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butCancel_Click (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butDots1_Click()
1         On Error GoTo ErrHandler
2         FilterValueDblClick ComboFilterBy1.Value, m_twb, TextBoxFilter1Value, m_LinesBook
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butDots1_Click (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butDots2_Click()
1         FilterValueDblClick ComboFilterBy2.Value, m_twb, TextBoxFilter2Value, m_LinesBook
2         Exit Sub
ErrHandler:
3         SomethingWentWrong "#butDots2_Click (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butDots3_Click()
1         On Error GoTo ErrHandler
2         CurrenciesDoubleClick
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butDots3_Click (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butOK_Click()
1         On Error GoTo ErrHandler
2         AddCayleyFiltersToMRU m_twb, ComboFilterBy1.Value, CastNumberStringToNumber(TextBoxFilter1Value.Value), _
                                ComboFilterBy2.Value, CastNumberStringToNumber(TextBoxFilter2Value.Value), Date
3         ButtonClicked = "OK"
4         HideForm Me    'See comments in method HideForm as to why we call it rather than simply Me.Hide
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butOK_Click (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub ComboFilterBy1_Change()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ComboFilterBy1_Change (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub ComboFilterBy2_Change()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ComboFilterBy2_Change (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxCurrencies_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
2         If Shift = 4 Then    'Alt key pressed
3             CurrenciesDoubleClick
4         End If
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#TextBoxCurrencies_KeyDown (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter1Value_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
2         If Shift = 4 Then    'Alt key pressed
3             FilterValueDblClick ComboFilterBy1.Value, m_twb, TextBoxFilter1Value, m_LinesBook
4         End If
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#TextBoxFilter1Value_KeyDown (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter1Value_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         FilterValueDblClick ComboFilterBy1.Value, m_twb, TextBoxFilter1Value, m_LinesBook
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBoxFilter1Value_MouseDown (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter2Value_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         FilterValueDblClick ComboFilterBy2.Value, m_twb, TextBoxFilter2Value, m_LinesBook
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBoxFilter2Value_MouseDown (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter2Value_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
2         If Shift = 4 Then    'Alt key pressed
3             FilterValueDblClick ComboFilterBy2.Value, m_twb, TextBoxFilter2Value, m_LinesBook
4         End If
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#TextBoxFilter2Value_KeyDown (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
1         On Error GoTo ErrHandler
2         ButtonClicked = "Cancel"
3         HideForm Me
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#UserForm_QueryClose (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Function CurrenciesDoubleClick()
          Dim Currencies
          Dim Res
1         On Error GoTo ErrHandler

2         Currencies = CurrenciesFromQuery("None", "None", "None", "None", False, 0, True, True, m_twb, m_twb.Worksheets(1), Date)

3         Res = ShowMultipleChoiceDialog(Currencies, sTokeniseString(TextBoxCurrencies.Value), _
                                         "Currencies to Include", "Only trades in the currencies you select" + vbLf + _
                                                                  "are retrieved from the trades workbook.", , Me, , , False)
4         If Not sArraysIdentical(Res, "#User Cancel!") Then
5             TextBoxCurrencies.Value = sConcatenateStrings(Res)
6         End If
7         Exit Function
ErrHandler:
8         SomethingWentWrong "#CurrenciesDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption

End Function

Private Sub TextBoxCurrencies_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         Cancel = True
3         CurrenciesDoubleClick
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TextBoxCurrencies_DblClick (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter1Value_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         Cancel = True
3         FilterValueDblClick ComboFilterBy1.Value, m_twb, TextBoxFilter1Value, m_LinesBook
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TextBoxFilter1Value_DblClick (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter2Value_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         Cancel = True
3         FilterValueDblClick ComboFilterBy2.Value, m_twb, TextBoxFilter2Value, m_LinesBook
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TextBoxFilter2Value_DblClick (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Function AlignDotsButton(TheTextBox As Control, TheButton As Control)
1         On Error GoTo ErrHandler
          Const Nudge = 1
2         TheButton.Height = TheTextBox.Height - 2 * Nudge
3         TheButton.Width = 14    'TheTextBox.Height
4         TheButton.Top = TheTextBox.Top + Nudge
5         TheButton.Left = TheTextBox.Left + TheTextBox.Width - TheButton.Width
6         TheButton.Visible = True
7         TheButton.TakeFocusOnClick = False
8         Exit Function
ErrHandler:
9         Throw "#AlignDotsButton (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshControls
' Author    : Philip Swannell
' Date      : 22-Dec-2016
' Purpose   : Housekeeping on the form - makes it wider if necessary, hides and unhides
'             the "dots" buttons and ensures they are correctly placed, disables some
'             controls when necessary.
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshControls()
1         On Error GoTo ErrHandler

          Dim RequiredWidth
          Dim TextArray(1 To 7, 1 To 1) As String
2         If m_ResizerInitialised Then
3             TextArray(1, 1) = TextBoxFile.Value
4             TextArray(2, 1) = ComboFilterBy1.Value
5             TextArray(3, 1) = TextBoxFilter1Value.Value
6             TextArray(4, 1) = ComboFilterBy2.Value
7             TextArray(5, 1) = TextBoxFilter2Value.Value
8             TextArray(6, 1) = ComboAssetClasses.Value
9             TextArray(7, 1) = TextBoxCurrencies.Value
10            RequiredWidth = sColumnMax(sStringWidth(TextArray, "Calibri", 11))(1, 1) + 40
11            If RequiredWidth > 800 Then RequiredWidth = 800
12            If RequiredWidth > TextBoxCurrencies.Width Then
13                m_clsResizer.ResizeControls Me, Me.Height, Me.Width, Me.Height, Me.Width + RequiredWidth - TextBoxCurrencies.Width
14            End If
15        End If

16        If Not Me.ActiveControl Is Nothing Then
17            Select Case Me.ActiveControl.Name
              Case TextBoxFilter1Value.Name, butDots1.Name
18                AlignDotsButton TextBoxFilter1Value, butDots1
19                If ComboFilterBy1.Value = "None" Then butDots1.Visible = False
20                butDots2.Visible = False
21                butDots3.Visible = False
22            Case TextBoxFilter2Value.Name, butDots2.Name
23                butDots1.Visible = False
24                AlignDotsButton TextBoxFilter2Value, butDots2
25                If ComboFilterBy2.Value = "None" Then butDots2.Visible = False
26                butDots3.Visible = False
27            Case TextBoxCurrencies.Name, butDots3.Name
28                butDots1.Visible = False
29                butDots2.Visible = False
30                AlignDotsButton TextBoxCurrencies, butDots3
31            Case Else
32                butDots1.Visible = False
33                butDots2.Visible = False
34                butDots3.Visible = False
35            End Select
36        End If
37        If ComboFilterBy1.Value = "None" Then
38            TextBoxFilter1Value.Value = ""
39            TextBoxFilter1Value.Enabled = False
40        Else
41            TextBoxFilter1Value.Enabled = True
42        End If
43        If ComboFilterBy2.Value = "None" Then
44            TextBoxFilter2Value.Value = ""
45            TextBoxFilter2Value.Enabled = False
46        Else
47            TextBoxFilter2Value.Enabled = True
48        End If
49        Exit Sub
ErrHandler:
50        Throw "#RefreshControls (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub ComboFilterBy1_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ComboFilterBy1_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter1Value_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBoxFilter1Value_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub ComboFilterBy2_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ComboFilterBy2_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxFilter2Value_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBoxFilter2Value_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub ComboAssetClasses_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ComboAssetClasses_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub TextBoxCurrencies_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBoxCurrencies_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub CheckBoxCompress_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#CheckBoxCompress_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butOK_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butOK_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butCancel_Enter()
1         On Error GoTo ErrHandler
2         RefreshControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butCancel_Enter (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butOK
3         UnHighlightFormControl butCancel
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butOK_MouseMove (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butCancel_MouseMove (line " & CStr(Erl) & "): " & Err.Description & "!", vbExclamation, Me.Caption
End Sub

Private Function CastNumberStringToNumber(TheInput As String)
1         On Error GoTo ErrHandler
2         CastNumberStringToNumber = CDbl(TheInput)
3         Exit Function
ErrHandler:
4         CastNumberStringToNumber = TheInput
End Function


