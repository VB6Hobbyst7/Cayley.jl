VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMultipleChoice 
   Caption         =   "Select"
   ClientHeight    =   7908
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   7896
   OleObjectBlob   =   "frmMultipleChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMultipleChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------------------------------------
' Module    : frmMultipleChoice
' Author    : Philip Swannell
' Date      : 22-Oct-2013
' Purpose   : Implement a dialog to allow a user to choose multiple values from a list _
'             of input strings or numbers. Wrapped by function ShowMultipleChoiceDialog
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private mclsResizer As clsFormResizer
Private m_AllowNoneLeftButton As Boolean
Private m_AllowNoneMiddleButton As Boolean
Public ReturnValue As Variant
Public m_ButtonClicked As String
Private m_TheList As Variant
Private m_SetHeightTo As Double
Private m_SuppressChangeEvent As Boolean
Private m_WithSort As Boolean

Const Max_Width_List As Long = 500
Const Max_Rows_List As Long = 40        '<-- Applies when CheckBoxes not shown. When checkboxes are shown, _
                                         each row is 13 points not 10 points tall so we ratio down the _
                                         maximum no of rows to show in the ListBox.
Const Min_Rows_List As Long = 10
Const v_gap As Long = 10
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Initialise
' Author    : Philip Swannell
' Date      : 22-Oct-2013
' Purpose   : Populating and sizing elements on the form
' -----------------------------------------------------------------------------------------------------------------------
Sub Initialise(TheList As Variant, _
        InitialChoices As Variant, _
        Title As String, _
        TopText As String, _
        ShowCheckBoxes As Boolean, _
        Caption1 As String, _
        Caption2 As String, _
        Caption3 As String, AllowNoneLeftButton, AllowNoneMiddleButton, WithSort As Boolean, CheckBoxCaption As String, CheckBoxValue As Boolean)

          Dim Accelerator As String
          Dim c As control
          Dim Have3Buttons As Boolean
          Dim i As Long
          Dim MatchRes As Variant
          Dim Min_Width_list As Double
          Dim NumRowsToShow As Long

1         On Error GoTo ErrHandler

2         For Each c In Me.Controls
3             If Not c Is lblSort Then
4                 c.Font.Name = "Segoe UI"
5                 If c Is LstBxChoices Then
6                     c.Font.Size = 8.5
7                 Else
8                     c.Font.Size = 9
9                 End If
10            End If
11        Next c

12        Me.caption = Title
13        butOK.caption = ProcessAmpersands(Caption1, Accelerator)
14        If Accelerator <> vbNullString Then butOK.Accelerator = Accelerator
15        butCancel.caption = ProcessAmpersands(Caption2, Accelerator)
16        If Accelerator <> vbNullString Then butCancel.Accelerator = Accelerator
17        If Caption3 = vbNullString Then
18            but3.Visible = False
19            butOK.Default = True
20            butOK.Cancel = False
21            butCancel.Cancel = True
22            butCancel.Default = False
23        Else
24            Have3Buttons = True
25            but3.caption = ProcessAmpersands(Caption3, Accelerator)
26            If Accelerator <> vbNullString Then but3.Accelerator = Accelerator
27            but3.Default = True
28            but3.Cancel = False
29            butOK.Default = False
30            butOK.Cancel = False
31            butCancel.Cancel = True
32            butCancel.Default = False
33        End If

34        m_TheList = TheList
35        m_AllowNoneLeftButton = AllowNoneLeftButton
36        m_AllowNoneMiddleButton = AllowNoneMiddleButton
37        m_WithSort = WithSort

38        LstBxChoices.List = TheList
39        LstBxChoices_Change        'disables the OK button if necessary

          'Horizontal postioning
          'Use lblAutoSize's AutoSize method to set the width of LstBxChoices (which does not have an autosize method)
40        lblAutoSize.caption = TopText + vbLf + sConcatenateStrings(TheList, vbLf)
41        lblAutoSize.Width = Max_Width_List + 100
42        lblAutoSize.AutoSize = True

43        LstBxChoices.Width = lblAutoSize.Width + 25        'The 25 allows for the vertical ScrollBar

44        lblAutoSize.caption = vbNullString
45        lblAutoSize.Visible = False

46        If Have3Buttons Then
              Const butGap = 5
47            but3.Left = butOK.Left + butOK.Width + butGap
48            butCancel.Left = but3.Left + but3.Width + butGap
49        Else
50            butCancel.Left = butOK.Left + butOK.Width + butGap
51        End If

52        Min_Width_list = butCancel.Left + butCancel.Width - butOK.Left
53        If LstBxChoices.Width > Max_Width_List Then
54            LstBxChoices.Width = Max_Width_List
55        ElseIf LstBxChoices.Width < Min_Width_list Then
56            LstBxChoices.Width = Min_Width_list
57        End If
58        If WithSort Then
59            SortControlSetDirection lblSort, sbDirectionFlat
60            lblSort.AutoSize = True
61            lblSort.AutoSize = False
62            lblSort.Left = LstBxChoices.Left
63            lblSort.Width = LstBxChoices.Width - 2
              Dim TheListSorted
64            TheListSorted = sSortedArray(TheList)
65            If sArraysIdentical(TheList, TheListSorted) Then
66                SortControlSetDirection lblSort, sbDirectionDown
67            ElseIf sArraysIdentical(TheList, sColumnReverse(TheListSorted)) Then
68                SortControlSetDirection lblSort, sbDirectionUp
69            Else
70                SortControlSetDirection lblSort, sbDirectionFlat
71            End If
72        End If

73        If TopText <> vbNullString Then
74            lblTopText.Width = LstBxChoices.Width
75            lblTopText.caption = TopText
76            lblTopText.AutoSize = False
77            lblTopText.AutoSize = True
78        Else
79            lblTopText.Visible = False
80        End If

81        If CheckBoxCaption <> vbNullString Then
82            With CheckBox1
83                .Value = CheckBoxValue
84                .caption = CheckBoxCaption
85                .Width = 1000
86                .AutoSize = True
87                .AutoSize = False
88                .Left = butOK.Left
89            End With
90        Else
91            CheckBox1.Visible = False
92        End If

93        If Not WithSort Then lblSort.Visible = False

94        Me.Width = LstBxChoices.Left + LstBxChoices.Width + 15
95        If CheckBoxCaption <> vbNullString Then
96            If Me.Width < CheckBox1.Left + CheckBox1.Width + Me.Width - Me.InsideWidth Then
97                Me.Width = CheckBox1.Left + CheckBox1.Width + Me.Width - Me.InsideWidth
98            End If
99        End If

          'Vertical positioning
100       If TopText <> vbNullString Then
101           lblTopText.Top = v_gap
102           If WithSort Then
103               lblSort.Top = lblTopText.Top + lblTopText.Height + v_gap
104               LstBxChoices.Top = lblSort.Top + lblSort.Height
105           Else
106               LstBxChoices.Top = lblTopText.Top + lblTopText.Height + v_gap
107           End If
108       Else
109           If WithSort Then
110               lblSort.Top = v_gap
111               LstBxChoices.Top = lblSort.Top + lblSort.Height
112           Else
113               LstBxChoices.Top = v_gap
114           End If
115       End If

116       NumRowsToShow = sNRows(TheList)
117       If NumRowsToShow < Min_Rows_List Then
118           NumRowsToShow = Min_Rows_List
119       ElseIf NumRowsToShow > CLng(Max_Rows_List * IIf(ShowCheckBoxes, 10, 13) / 13) Then
120           NumRowsToShow = CLng(Max_Rows_List * IIf(ShowCheckBoxes, 10, 13) / 13)
121       End If
122       m_SetHeightTo = NumRowsToShow * IIf(ShowCheckBoxes, 13, 10) + 2
123       LstBxChoices.Height = m_SetHeightTo
124       LstBxChoices.IntegralHeight = False
125       LstBxChoices.IntegralHeight = True
126       butOK.Top = LstBxChoices.Top + LstBxChoices.Height + v_gap
127       butCancel.Top = butOK.Top
128       but3.Top = butOK.Top

129       LstBxChoices.MultiSelect = fmMultiSelectExtended

130       CheckBox1.Top = butOK.Top + butOK.Height + v_gap

131       With lblHint
132           .caption = "Keyboard Shortcuts"
133           .Font.Underline = True
134           .Width = 500
135           .AutoSize = False
136           .AutoSize = True
137           If CheckBox1.Visible = False Then
138               .Top = butOK.Top + butOK.Height + v_gap
139           Else
140               .Top = CheckBox1.Top + CheckBox1.Height + v_gap
141           End If

142       End With

          'Me.Height-Me.InsideHeight is a measure of the height of the title bar at the top of the form, which changes in one version of Excel to the next...
143       Me.Height = lblHint.Top + lblHint.Height + v_gap + (Me.Height - Me.InsideHeight + 5)

          'Set the tags needed by the resizer
144       LstBxChoices.Tag = "HW"
145       butOK.Tag = "T"
146       butCancel.Tag = "T"
147       but3.Tag = "T"
148       lblHint.Tag = "T"
149       If WithSort Then lblSort.Tag = "W"
150       If CheckBoxCaption <> vbNullString Then CheckBox1.Tag = "T"

151       If ShowCheckBoxes Then LstBxChoices.ListStyle = fmListStyleOption

          'Initialise which elements are selected
152       MatchRes = sMatch(TheList, InitialChoices)
153       Force2DArray MatchRes
          Dim HaveSetIndex As Boolean
154       For i = 1 To sNRows(MatchRes)
155           If VarType(MatchRes(i, 1)) <> vbString Then
156               If Not HaveSetIndex Then
157                   LstBxChoices.ListIndex = i - 1
158                   HaveSetIndex = True
159               End If
160               LstBxChoices.Selected(i - 1) = True
161           End If
162       Next i

          'Create the instance of the class
163       CreateFormResizer mclsResizer
          Dim MinimumHeight As Double
          Dim MinimumWidth As Double
164       MinimumWidth = butCancel.Left + butCancel.Width + 22        'leave space for the grab handle
165       MinimumWidth = SafeMax(MinimumWidth, lblHint.Left + lblHint.Width + Me.Width - Me.InsideWidth + 5)
166       MinimumHeight = Me.Height - LstBxChoices.Height + IIf(ShowCheckBoxes, 63, 53)        'allow 5 rows to show

          'Tell it which form it's handling
167       mclsResizer.Initialise Me, MinimumHeight, MinimumWidth

168       Exit Sub
ErrHandler:
169       Throw "#Initialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butCancel_Click()
1         On Error GoTo ErrHandler

2         Me.ReturnValue = "#User Cancel!"
3         Me.m_ButtonClicked = butCancel.caption
4         HideForm Me

5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butCancel_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butOK_Click()
          Dim ChooseVector As Variant
          Dim i As Long
          Dim Res As Variant

1         On Error GoTo ErrHandler

2         ChooseVector = sReshape(False, Me.LstBxChoices.ListCount, 1)
3         For i = 1 To Me.LstBxChoices.ListCount
4             ChooseVector(i, 1) = Me.LstBxChoices.Selected(i - 1)
5         Next i
6         If sColumnOr(ChooseVector)(1, 1) Then
7             Res = sMChoose(m_TheList, ChooseVector)
8             Force2DArray Res
9             Me.ReturnValue = Res
10        Else
11            Me.ReturnValue = Empty
12        End If
13        Me.m_ButtonClicked = butOK.caption

14        HideForm Me

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#butOK_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub but3_Click()
          Dim ChooseVector As Variant
          Dim i As Long
          Dim Res As Variant

1         On Error GoTo ErrHandler

2         ChooseVector = sReshape(False, Me.LstBxChoices.ListCount, 1)
3         For i = 1 To Me.LstBxChoices.ListCount
4             ChooseVector(i, 1) = Me.LstBxChoices.Selected(i - 1)
5         Next i
6         If sColumnOr(ChooseVector)(1, 1) Then
7             Res = sMChoose(m_TheList, ChooseVector)
8             Force2DArray Res
9             Me.ReturnValue = Res
10        Else
11            Me.ReturnValue = Empty
12        End If
13        Me.m_ButtonClicked = but3.caption

14        HideForm Me

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#but3_Click (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub SelectAll(SwitchOn As Boolean)
          Dim i As Long
1         On Error GoTo ErrHandler
2         For i = 0 To LstBxChoices.ListCount - 1
3             LstBxChoices.Selected(i) = SwitchOn
4         Next i
5         Exit Sub
ErrHandler:
6         Throw "#SelectAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub lblHint_Click()
          Dim Prompt

1         On Error GoTo ErrHandler
2         Prompt = sArrayStack("Select All", "Ctrl A", _
              "Select None", "Ctrl N", _
              "Invert Selection", "Ctrl I", _
              "Multi-select", "Ctrl-Click or Shift-Click", _
              vbNullString, vbNullString, _
              "Nudge Selected Up", "Alt Up", _
              "Move Selected to Top", "Alt Home", _
              "Nudge Selected Down", "Alt Down", _
              "Move Selected to Bottom", "Alt End", _
              vbNullString, vbNullString, _
              "Sort", "Ctrl S")

3         Prompt = sReshape(Prompt, sNRows(Prompt) / 2, 2)
4         Prompt = sJustifyArrayOfStrings(Prompt, "Segoe UI", 9, vbTab)
5         Prompt = sConcatenateStrings(Prompt, vbLf)

6         MsgBoxPlus Prompt, , "Keyboard Shortcuts", , , , , , , , , , Me

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#lblHint_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub lblSort_Click()
1         On Error GoTo ErrHandler
2         Sort

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#lblSort_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub LstBxChoices_Change()
          Dim AnySelected As Boolean
          Dim i As Long
1         On Error GoTo ErrHandler
2         If m_SuppressChangeEvent Then Exit Sub
3         If (Not m_AllowNoneLeftButton) Or (Not m_AllowNoneMiddleButton) Then
4             For i = 0 To LstBxChoices.ListCount
5                 If LstBxChoices.Selected(i) Then
6                     AnySelected = True
7                     Exit For
8                 End If
9             Next i
10            butOK.Enabled = AnySelected Or m_AllowNoneLeftButton
11            but3.Enabled = AnySelected Or m_AllowNoneMiddleButton
12        End If
13        Exit Sub
ErrHandler:
14        Throw "#LstBxChoices_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LstBxChoices_KeyDown
' Author    : Philip Swannell
' Date      : 22-Oct-2013
' Purpose   : Make Ctrl+A and Ctrl+N do useful things - select all or select none, plus more added over time
' -----------------------------------------------------------------------------------------------------------------------
Private Sub LstBxChoices_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
          Dim i As Long

2         If KeyCode = 65 And Shift = 2 Then        'Ctrl A
3             SelectAll True
4         ElseIf KeyCode = 78 And Shift = 2 Then        'Ctrl N
5             SelectAll False
6         ElseIf KeyCode = 73 And Shift = 2 Then        'Ctrl I
7             For i = 0 To LstBxChoices.ListCount - 1
8                 LstBxChoices.Selected(i) = Not (LstBxChoices.Selected(i))
9             Next i
10        ElseIf KeyCode = 38 And Shift = 4 Then        'Alt Up
11            PromoteSelected -1
12        ElseIf KeyCode = 36 And Shift = 4 Then        'Alt Home
13            PromoteSelected -LstBxChoices.ListCount
14        ElseIf KeyCode = 40 And Shift = 4 Then        'Alt Down
15            PromoteSelected 1
16        ElseIf KeyCode = 35 And Shift = 4 Then        'Alt End
17            PromoteSelected LstBxChoices.ListCount
18        ElseIf KeyCode = 83 And Shift = 2 Then        'Ctrl S
19            Sort
20        End If
21        Exit Sub
ErrHandler:
22        SomethingWentWrong "#LstBxChoices_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Sort
' Author     : Philip Swannell
' Date       : 03-May-2018
' Purpose    : Sort the elements of LstBoxChoices
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub Sort()
          Dim i As Long
          Dim NewList
          Dim NewListAscending
          Dim NewListDescending
          Dim NewSelected
          Dim OldHeight As Double
          Dim OldList
          Dim OldListIndex As Long
          Dim OldTopIndex As Long
          Dim Selected
          Dim SortAscending As Boolean
          Dim SortedIntegers

1         On Error GoTo ErrHandler

2         m_SuppressChangeEvent = True
3         OldHeight = LstBxChoices.Height
4         OldTopIndex = LstBxChoices.TopIndex
5         OldListIndex = LstBxChoices.ListIndex

6         OldList = To1Based2D(LstBxChoices.List)
7         Selected = sReshape(False, LstBxChoices.ListCount, 1)
8         For i = 1 To LstBxChoices.ListCount
9             If LstBxChoices.Selected(i - 1) Then Selected(i, 1) = True
10        Next i

11        NewListAscending = sArrayRange(OldList, Selected, sIntegers(LstBxChoices.ListCount))
12        NewListAscending = sSortedArray(NewListAscending)
13        NewListDescending = sColumnReverse(NewListAscending)
14        SortAscending = Not sArraysIdentical(OldList, sSubArray(NewListAscending, 1, 1, , 1))
15        If SortAscending Then
16            NewList = sSubArray(NewListAscending, 1, 1, , 1)
17            NewSelected = sSubArray(NewListAscending, 1, 2, , 1)
18            SortedIntegers = sSubArray(NewListAscending, 1, 3, , 1)
19            SortControlSetDirection lblSort, sbDirectionDown
20        Else
21            NewList = sSubArray(NewListDescending, 1, 1, , 1)
22            NewSelected = sSubArray(NewListDescending, 1, 2, , 1)
23            SortedIntegers = sSubArray(NewListDescending, 1, 3, , 1)
24            SortControlSetDirection lblSort, sbDirectionUp
25        End If

26        m_TheList = NewList
27        LstBxChoices.List = NewList

28        LstBxChoices.IntegralHeight = False
29        LstBxChoices.Height = OldHeight

30        For i = 1 To LstBxChoices.ListCount
31            LstBxChoices.Selected(i - 1) = NewSelected(i, 1)
32        Next i

33        LstBxChoices.TopIndex = OldTopIndex
34        LstBxChoices.ListIndex = sMatch(OldListIndex + 1, SortedIntegers) - 1

35        m_SuppressChangeEvent = False

36        Exit Sub
ErrHandler:
37        Throw "#Sort (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub PromoteSelected(Offset As Long)
          Dim i As Long
          Dim NewIntegers As Variant
          Dim NewList
          Dim NewTopIndex
          Dim OldHeight
          Dim OldListIndex As Long
          Dim PromoteRes
          Dim Selected

1         On Error GoTo ErrHandler
2         m_SuppressChangeEvent = True
3         OldListIndex = LstBxChoices.ListIndex

4         NewTopIndex = LstBxChoices.TopIndex + Offset
5         NewTopIndex = SafeMax(0, NewTopIndex)
6         NewTopIndex = SafeMin(LstBxChoices.ListCount - 1, NewTopIndex)

7         OldHeight = LstBxChoices.Height

8         Selected = sReshape(False, LstBxChoices.ListCount, 1)
9         For i = 1 To LstBxChoices.ListCount
10            If LstBxChoices.Selected(i - 1) Then Selected(i, 1) = True
11        Next i
12        PromoteRes = sPromote(sArrayRange(To1Based2D(LstBxChoices.List), sIntegers(LstBxChoices.ListCount)), Selected, Offset)
13        NewList = sSubArray(PromoteRes, 1, 1, , 1)
14        NewIntegers = sSubArray(PromoteRes, 1, 2, , 1)
15        m_TheList = NewList
16        LstBxChoices.List = NewList

17        LstBxChoices.IntegralHeight = False
18        LstBxChoices.Height = OldHeight

19        For i = 1 To LstBxChoices.ListCount
20            LstBxChoices.Selected(i - 1) = Selected(i, 1)
21        Next i

22        LstBxChoices.TopIndex = NewTopIndex

23        LstBxChoices.ListIndex = sMatch(OldListIndex + 1, NewIntegers) - 1

24        SortControlSetDirection lblSort, sbDirectionFlat
25        m_SuppressChangeEvent = False
26        Exit Sub

27        Exit Sub
ErrHandler:
28        Throw "#PromoteSelected (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
1         On Error GoTo ErrHandler

2         Cancel = True
3         butCancel_Click

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#UserForm_QueryClose (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub LstBxChoices_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         UserForm_MouseMove Button, Shift, x, y
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         UnHighlightFormControl butCancel
3         UnHighlightFormControl but3
4         UnHighlightFormControl butOK
5         UnHighlightFormControl lblSort

6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         HighlightFormControl butOK
3         UnHighlightFormControl but3
4         UnHighlightFormControl butCancel
5         UnHighlightFormControl lblSort

6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#butOK_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub
Private Sub but3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         HighlightFormControl but3
3         UnHighlightFormControl butOK
4         UnHighlightFormControl butCancel
5         UnHighlightFormControl lblSort

6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#but3_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         HighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but3
5         UnHighlightFormControl lblSort

6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#butCancel_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub lblSort_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

1         On Error GoTo ErrHandler
2         UnHighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but3
5         HighlightFormControl lblSort
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#lblSort_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortControlSetDirection - adapted from SortButtonSetDirection
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Sets the text on a sort button. Needs Wingdings and Wingdings 3 to be installed
'             for nice-looking symbols on buttons. If they're not installed we do as best we can
'             with "-", "v" and "^" characters.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SortControlSetDirection(b As control, Direction As EnmsbDirection)

1         On Error GoTo ErrHandler
          Const FontSize = 10

2         b.Font.Bold = False

3         Select Case Direction
              Case sbDirectionFlat
4                 If sFontIsInstalled("Wingdings") Then
5                     b.caption = Chr$(108)
6                     b.Font.Name = "Wingdings"
7                 Else
8                     b.caption = "-"
9                     b.Font.Name = "Arial"
10                End If
11                b.Font.Size = FontSize
12            Case sbDirectionDown
13                If sFontIsInstalled("Wingdings 3") Then
14                    b.caption = Chr$(112)
15                    b.Font.Name = "Wingdings 3"
16                Else
17                    b.caption = "v"
18                    b.Font.Name = "Arial"
19                    b.Font.Bold = True
20                End If
21                b.Font.Size = FontSize
22            Case sbDirectionUp
23                If sFontIsInstalled("Wingdings 3") Then
24                    b.caption = Chr$(113)
25                    b.Font.Name = "Wingdings 3"
26                Else
27                    b.caption = "^"
28                    b.Font.Name = "Arial"
29                    b.Font.Bold = True
30                End If
31                b.Font.Size = FontSize
32        End Select
33        Exit Sub
ErrHandler:
34        Throw "#SortControlSetDirection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
