VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegularExpression 
   Caption         =   "UserForm1"
   ClientHeight    =   8676.001
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8778.001
   OleObjectBlob   =   "frmRegularExpression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegularExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mSuggestedStrings As Variant
Public ReturnArray
Public mclsResizer As clsFormResizer
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Initialise
' Author    : Philip Swannell
' Date      : 02-May-2016
' Purpose   : Initialise the dialog.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Initialise(Title As String, ActionText As String, AttributeName As String, Optional SuggestedStrings, Optional InitialisationArray As Variant)

          Const ValidOperators = ",contains,equals,does not equal,begins with,does not begin with,ends with,does not end with,does not contain"

1         On Error GoTo ErrHandler

2         Me.caption = Title
3         Me.Label1.caption = ActionText
4         Me.LeftComboBox1.List = sTokeniseString(ValidOperators + ",in,not in")
5         Me.LeftComboBox2.List = sTokeniseString(ValidOperators)
6         If Not IsMissing(SuggestedStrings) Then
7             Me.RightComboBox1.List = SuggestedStrings
8             Me.RightComboBox2.List = SuggestedStrings
9             Me.mSuggestedStrings = SuggestedStrings
10        End If
11        Me.LeftComboBox1.ListIndex = 1
12        Me.OptAnd1.Value = False
13        Me.OptOr1.Value = True

14        If Not IsMissing(InitialisationArray) Then
15            If Not IsEmpty(InitialisationArray) Then
16                LeftComboBox1.Value = InitialisationArray(1, 2)
17                If InitialisationArray(1, 2) = "in" Or InitialisationArray(1, 2) = "not in" Then
18                    TextBox1.Value = InitialisationArray(1, 3)
19                Else
20                    RightComboBox1.Value = InitialisationArray(1, 3)
21                End If
22                If sNRows(InitialisationArray) > 1 Then
23                    LeftComboBox2.Value = InitialisationArray(2, 2)
24                    RightComboBox2.Value = InitialisationArray(2, 3)
25                    Me.OptAnd1.Value = InitialisationArray(2, 1) = "AND"
26                    Me.OptOr1.Value = InitialisationArray(2, 1) = "OR"
27                End If
28            End If
29        End If

30        Me.Label2.caption = AttributeName
31        If Me.RightComboBox1.Visible Then
32            Me.RightComboBox1.SetFocus
33        Else
34            Me.TextBox1.SetFocus
35        End If

36        Me.ResizeControls
37        Exit Sub
ErrHandler:
38        Throw "#Initialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeControls
' Author    : Philip Swannell
' Date      : 02-May-2016
' Purpose   : Set the layout of the form.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub ResizeControls()
          Const FontName = "Segoe UI"
          Const FontSize = 10
          Dim butHeight As Double
          Dim butWidth As Double
          Dim LeftMargin As Double
          Dim TopMargin As Double
          Dim vNudge1 As Double

          Dim ScaleFactor As Variant

1         On Error GoTo ErrHandler
2         ScaleFactor = sStringWidth("abcdefghijclmnopqrstuvwxyz", FontName, FontSize)(1, 1) / 135.75

3         vNudge1 = 5 * ScaleFactor
4         LeftMargin = 10 * ScaleFactor
5         TopMargin = 10 * ScaleFactor
6         butHeight = 36 * ScaleFactor
7         butWidth = 78 * ScaleFactor

8         With Label1
9             .Top = TopMargin
10            .Left = LeftMargin
11            .Font.Name = FontName
12            .Font.Size = FontSize
13            .Width = 1000
14            .AutoSize = True
15            .AutoSize = False
16        End With

17        With Label2
18            .Left = LeftMargin
19            .Top = Label1.Top + Label1.Height + 5 * ScaleFactor
20            .Font.Name = FontName
21            .Font.Size = FontSize
22            .Width = 1000
23            .AutoSize = True
24            .AutoSize = False
25        End With

26        With LeftComboBox1
27            .Left = LeftMargin
28            .Top = Label2.Top + Label2.Height + 3 * ScaleFactor
29            .Font.Name = FontName
30            .Font.Size = FontSize
31            .AutoSize = True
32            .AutoSize = False
33            .Width = sMaxOfArray(sStringWidth(sSubArray(.List, 1, 1), FontName, FontSize)) + 35 * ScaleFactor
34        End With

35        PlaceGuardLabels LeftComboBox1, Label3, Label4

36        With OptAnd1
37            .Left = LeftMargin * 2
38            .Font.Name = FontName
39            .Font.Size = FontSize
40            .AutoSize = True
41            .AutoSize = True
42            .Top = LeftComboBox1.Top + LeftComboBox1.Height + vNudge1
43        End With

44        With OptOr1
45            .Font.Name = FontName
46            .Font.Size = FontSize
47            .AutoSize = True
48            .AutoSize = True
49            .Left = OptAnd1.Left + OptAnd1.Width
50            .Top = OptAnd1.Top
51        End With

52        With LeftComboBox2
53            .Top = OptAnd1.Top + OptAnd1.Height + vNudge1
54            .Font.Name = FontName
55            .Font.Size = FontSize
56            .Height = LeftComboBox1.Height
57            .Width = LeftComboBox1.Width
58            .Left = LeftComboBox1.Left
59        End With

60        PlaceGuardLabels LeftComboBox2, Label5, Label6

61        With RightComboBox1
62            .Left = LeftComboBox1.Left + LeftComboBox1.Width + LeftMargin
63            .Font.Name = FontName
64            .Font.Size = FontSize
65            .Top = LeftComboBox1.Top
66            .Height = LeftComboBox1.Height
67            .Width = LeftComboBox1.Width * 2
68        End With

69        PlaceGuardLabels RightComboBox1, Label7, Label8

70        With RightComboBox2
71            .Font.Name = FontName
72            .Font.Size = FontSize
73            .Top = LeftComboBox2.Top
74            .Height = LeftComboBox2.Height
75            .Width = LeftComboBox2.Width * 2
76            .Left = RightComboBox1.Left
77        End With

78        PlaceGuardLabels RightComboBox2, Label9, Label10

79        With TextBox1
80            .Font.Name = FontName
81            .Font.Size = FontSize
82            .Top = RightComboBox1.Top
83            .Left = RightComboBox1.Left
84            .Height = RightComboBox1.Height
85            .Width = RightComboBox1.Width
86        End With

87        With butList
88            .Font.Name = FontName
89            .Font.Size = FontSize
90            .Height = butHeight * 2 / 3
91            .Width = butWidth * 2 / 3
92            .Left = RightComboBox1.Left + RightComboBox1.Width - .Width
93            .Top = RightComboBox1.Top - .Height - 5
94        End With

95        With butCancel
96            .Font.Name = FontName
97            .Font.Size = FontSize
98            .Height = butHeight
99            .Width = butWidth
100           .Left = RightComboBox1.Left + RightComboBox1.Width - .Width
101           .Top = RightComboBox2.Top + RightComboBox2.Height + 10
102       End With

103       With butOK
104           .Font.Name = FontName
105           .Font.Size = FontSize
106           .Height = butHeight
107           .Width = butWidth
108           .Left = butCancel.Left - .Width - 10
109           .Top = butCancel.Top
110       End With

111       Me.Width = RightComboBox1.Left + RightComboBox1.Width + LeftMargin + Me.Width - Me.InsideWidth
112       Me.Height = butOK.Top + butOK.Height + LeftMargin + Me.Height - Me.InsideHeight + 8        'need the extra 8 for the form resizer grab handle

          'Set up resizer...
113       RightComboBox1.Tag = "W"
114       RightComboBox2.Tag = "W"
115       TextBox1.Tag = "W"
116       Label7.Tag = "W"        'Not sure if it's necessary to resize the "guard labels"
117       Label8.Tag = "W"
118       Label9.Tag = "W"
119       Label10.Tag = "W"
120       butList.Tag = "L"
121       butOK.Tag = "L"
122       butCancel.Tag = "L"
123       CreateFormResizer mclsResizer

124       mclsResizer.Initialise Me, Me.Height, Me.Width, Me.BackColor

125       Exit Sub
ErrHandler:
126       Throw "#ResizeControls (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PlaceGuardLabels
' Author    : Philip Swannell
' Date      : 03-May-2016
' Purpose   : The behaviour of a combo box appears to depend on the controls that surround
'             in on the form! If there is another combobox just beneath it then hitting the
'             down key when the last item in the combo boxes dropdown list is selected switches
'             focus to the combo box beneath. This is surely a bug. The work-around is to place
'             zero-height labels above and below the combo box.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub PlaceGuardLabels(cb As control, L1 As control, L2 As control)
1         On Error GoTo ErrHandler
2         With L1
3             .Left = cb.Left
4             .Width = cb.Width
5             .Top = cb.Top - 2
6             .caption = vbNullString
7             .Height = 0
8         End With
9         With L2
10            .Left = cb.Left
11            .Width = cb.Width
12            .Top = cb.Top + cb.Height + 2
13            .caption = vbNullString
14            .Height = 0
15        End With
16        Exit Sub
ErrHandler:
17        Throw "#PlaceGuardLabels (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OKAllowed
' Author    : Philip Swannell
' Date      : 02-May-2016
' Purpose   : Encapsulate whether the user has entered data that "makes sense" if not we
'             disble the OK button.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OKAllowed() As Boolean

1         On Error GoTo ErrHandler
2         If Me.LeftComboBox1.Value = "in" Or Me.LeftComboBox1.Value = "not in" Then
3             OKAllowed = Len(TextBox1.Value) > 0
4             Exit Function
5         End If

6         OKAllowed = True
7         If LeftComboBox1.Value = vbNullString Then
8             OKAllowed = False
9             Exit Function
10        End If
11        If LeftComboBox1.Value <> "equals" And LeftComboBox1.Value <> "does not equal" Then
12            If RightComboBox1.Value = vbNullString Then
13                OKAllowed = False
14                Exit Function
15            End If
16        End If
17        If LeftComboBox2.Value = vbNullString And RightComboBox2.Value <> vbNullString Then
18            OKAllowed = False
19            Exit Function
20        End If
21        If LeftComboBox2.Value <> "equals" And LeftComboBox2.Value <> "does not equal" And LeftComboBox2.Value <> vbNullString Then
22            If RightComboBox2.Value = vbNullString Then
23                OKAllowed = False
24                Exit Function
25            End If
26        End If
27        Exit Function
ErrHandler:
28        Throw "#OKAllowed (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : EnableDisableControls
' Author    : Philip Swannell
' Date      : 30-Apr-2016
' Purpose   : When the top left combo box reads "in" or "not in", many of the other controls
'             need to be disabled. The top right combo box is replaced with a text box and a
'             button butList becomes visible. This method also disables the OK button if the current
'             inputs don't make sense.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub EnableDisableControls()

1         On Error GoTo ErrHandler
          Const Col_Grey = &H8000000F
          Const Col_White = &HFFFFFF
          Dim inMode As Boolean
          Dim TheColor As Double

2         Select Case Me.LeftComboBox1.Value
              Case "in", "not in"
3                 inMode = True
4                 TheColor = Col_Grey
5             Case Else
6                 inMode = False
7                 TheColor = Col_White
8         End Select

9         If Not OKAllowed() Then
              'We don't disable the control because _
               disabling the Default control has consequences for what happens when _
               the user hits Enter in other controls.  Hitting Enter becomes the same _
               as hitting Tab, which is not the user's intention (at least not mine). _
               So instead we just make the button look disabled.
10            butOK.ForeColor = &H80000011
11            butOK.TakeFocusOnClick = False
12            butOK.TabStop = False
13        Else
14            butOK.ForeColor = &H80000012
15            butOK.TakeFocusOnClick = True
16            butOK.TabStop = True
17        End If

18        butList.Visible = inMode And Not IsEmpty(mSuggestedStrings)
19        With LeftComboBox2
20            .Enabled = Not inMode
21            .TabStop = Not inMode
22            .BackColor = TheColor
23            If inMode Then .Value = vbNullString
24        End With
25        OptAnd1.Enabled = Not inMode

26        OptOr1.Enabled = Not inMode

27        With RightComboBox1
28            .TabStop = Not inMode
29            .Visible = Not inMode
30        End With

31        With RightComboBox2
32            .BackColor = TheColor
33            .Enabled = Not inMode
34            If inMode Then .Value = vbNullString
35        End With

36        With TextBox1
37            TextBox1.TabStop = inMode
38            TextBox1.Visible = inMode
39        End With

40        Exit Sub

ErrHandler:
41        Throw "#EnableDisableControls(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub LeftComboBox1_Change()
1         On Error GoTo ErrHandler
2         EnableDisableControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#LeftComboBox1_Change(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub LeftComboBox2_Change()
1         On Error GoTo ErrHandler
2         EnableDisableControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#LeftComboBox2_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_Change()
1         On Error GoTo ErrHandler
2         EnableDisableControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBox1_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub RightComboBox1_Change()
1         On Error GoTo ErrHandler
2         EnableDisableControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#RightComboBox1_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub RightComboBox2_Change()
1         On Error GoTo ErrHandler
2         EnableDisableControls
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#RightComboBox2_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RightComboBox1_Exit
' Author    : Philip Swannell
' Date      : 01-May-2016
' Purpose   : Keep the contents of TextBox1 in synch with the contents of RightComboBox1
'             TextBox1 is visible to the user only when LeftComboBox1 reads "in" or "not in"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RightComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

1         On Error GoTo ErrHandler
2         TextBox1.Value = RightComboBox1.Value
3         Exit Sub

ErrHandler:
4         SomethingWentWrong "#RightComboBox1_Exit(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         RightComboBox1.Value = TextBox1.Value
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBox1_Exit(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_Click()

1         On Error GoTo ErrHandler

          Dim TheArray As Variant

2         If Not OKAllowed() Then Exit Sub

3         TheArray = sReshape(vbNullString, 2, 3)

4         If Me.OptAnd1.Value = True Then
5             TheArray(2, 1) = "AND"
6         ElseIf Me.OptOr1.Value = True Then
7             TheArray(2, 1) = "OR"
8         End If

9         TheArray(1, 2) = Me.LeftComboBox1.Value
10        Select Case LeftComboBox1.Value
              Case "in", "not in"
11                TheArray(1, 3) = Me.TextBox1.Value
12            Case Else
13                TheArray(1, 3) = Me.RightComboBox1.Value
14        End Select

15        If Me.LeftComboBox2.Value <> vbNullString Then
16            TheArray(2, 2) = Me.LeftComboBox2.Value
17            TheArray(2, 3) = Me.RightComboBox2.Value
18        Else
19            TheArray = sSubArray(TheArray, 1, 1, 1)
20        End If

21        If sIsErrorString(TheArray) Then
22            MsgBoxPlus CStr(TheArray), vbCritical, "Regular Expression"
23        Else
24            ReturnArray = TheArray
25            HideForm Me
26        End If

27        Exit Sub

ErrHandler:
28        SomethingWentWrong "#butOK_Click(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butCancel_Click()

1         On Error GoTo ErrHandler

2         ReturnArray = "#User Cancel!"
3         HideForm Me

4         Exit Sub

ErrHandler:
5         SomethingWentWrong "#butCancel_Click(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butList_Click()

1         On Error GoTo ErrHandler
          Dim Res
2         If Not IsEmpty(mSuggestedStrings) Then
3             Res = ShowMultipleChoiceDialog(mSuggestedStrings, sTokeniseString(TextBox1.Value), "Select " & Label2.caption & "(s)", , , Me)
4             If Not sArraysIdentical(Res, "#User Cancel!") Then
5                 If IsEmpty(Res) Then
6                     TextBox1.Value = vbNullString
7                 Else
8                     TextBox1.Value = sConcatenateStrings(Res, ",")
9                 End If
10            End If
11            TextBox1.SetFocus
12        End If

13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#butList_Click(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         butList_Click
3         Exit Sub
ErrHandler:
4         Throw "#TextBox1_DblClick(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub LeftComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler
2         If KeyCode = 38 Then        'Up Key
3             If Me.LeftComboBox1.ListIndex = 0 Then
4                 Me.LeftComboBox1.SetFocus
5             End If
6         ElseIf KeyCode = 40 Then        'Down Key
7             LeftComboBox1.DropDown
8         ElseIf KeyCode = 39 Then        'Right Key
9             Select Case LCase$(LeftComboBox1.Value)
                  Case "in", "not in"
10                    Me.TextBox1.SetFocus
11                Case Else
12                    Me.RightComboBox1.SetFocus
13            End Select
14        End If

15        Exit Sub

ErrHandler:
16        Throw "#ComboBox1_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub LeftComboBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler

2         If KeyCode = 38 Then        'Up
3             If Me.LeftComboBox2.ListIndex = 0 Then
4                 Me.LeftComboBox2.SetFocus
5             End If
6         ElseIf KeyCode = 40 Then        'down
7             LeftComboBox2.DropDown
8         ElseIf KeyCode = 37 Then        'left
9             Me.RightComboBox1.SetFocus
10        ElseIf KeyCode = 39 Then        'Right
11            If LeftComboBox2.SelStart = Len(LeftComboBox2.Value) Then
12                Me.RightComboBox2.SetFocus
13            End If
14        End If

15        Exit Sub

ErrHandler:
16        Throw "#LeftComboBox2_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub RightComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 38 Then        'Up Key
3             If Me.RightComboBox1.ListIndex = 0 Then
4                 Me.RightComboBox1.SetFocus
5             End If
6         ElseIf KeyCode = 40 Then        'Down Key
7             RightComboBox1.DropDown
8         ElseIf KeyCode = 37 Then        'Left Key
9             If RightComboBox1.SelStart = 0 Then
10                Me.LeftComboBox1.SetFocus
11            End If
12        ElseIf KeyCode = 39 Then        'Right Key
13            If RightComboBox1.SelStart = Len(RightComboBox1.text) Then
14                If Me.LeftComboBox2.Enabled Then
15                    Me.LeftComboBox2.SetFocus
16                End If
17            End If
18        End If

19        Exit Sub

ErrHandler:
20        Throw "#RightComboBox1_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub RightComboBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 38 Then        'Up
3             If Me.RightComboBox2.ListIndex = 0 Then
4                 Me.RightComboBox2.SetFocus
5             End If
6         ElseIf KeyCode = 40 Then        'Down
7             If Me.RightComboBox2.ListIndex = Me.RightComboBox2.ListCount - 1 Then
8                 Me.RightComboBox2.SetFocus
9             End If

10            RightComboBox2.DropDown

11        ElseIf KeyCode = 37 Then        'Left
12            If RightComboBox2.SelStart = 0 Then
13                Me.LeftComboBox2.SetFocus
14            End If
15        End If

16        Exit Sub

ErrHandler:
17        Throw "#RightComboBox2_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub OptOr1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler

2         Select Case KeyCode
              Case 37 To 40        'Arrow keys
3                 OptAnd1.SetFocus
4                 OptAnd1.Value = True
5         End Select

6         Exit Sub

ErrHandler:
7         Throw "#OptOr1_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub OptAnd1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler

2         Select Case KeyCode
              Case 37 To 40        'Arrow keys
3                 OptOr1.SetFocus
4                 OptOr1.Value = True
5         End Select

6         Exit Sub

ErrHandler:
7         Throw "#OptAnd1_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler

2         If KeyCode = 40 Then        'down
3             butList_Click
4         ElseIf KeyCode = 81 And Shift = 4 Then
5             butList_Click

6         ElseIf KeyCode = 37 Then        'left
7             If TextBox1.SelStart = 0 Then
8                 LeftComboBox1.SetFocus
9             End If
10        ElseIf KeyCode = 39 Then        'Right
11            If TextBox1.SelStart = Len(TextBox1.text) Then
12                If LeftComboBox2.Enabled Then
13                    LeftComboBox2.SetFocus
14                End If
15            End If
16        End If

17        Exit Sub

ErrHandler:
18        Throw "#TextBox1_KeyDown(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

1         On Error GoTo ErrHandler

2         If (CloseMode = vbFormControlMenu) Then
3             Cancel = CLng(True)
4             ReturnArray = "#User Cancel!"
5             HideForm Me
6         End If

7         Exit Sub

ErrHandler:
8         Throw "#UserForm_QueryClose(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         If OKAllowed() Then
3             HighlightFormControl butOK
4         End If
5         UnHighlightFormControl butCancel
6         UnHighlightFormControl butList
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#butOK_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl butOK
3         HighlightFormControl butCancel
4         UnHighlightFormControl butList
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butCancel_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl butOK
3         UnHighlightFormControl butCancel
4         HighlightFormControl butList
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butList_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

1         On Error GoTo ErrHandler
2         UnHighlightFormControl butOK
3         UnHighlightFormControl butCancel
4         UnHighlightFormControl butList
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
