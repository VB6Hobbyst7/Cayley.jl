VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputBoxPlus 
   Caption         =   "Input"
   ClientHeight    =   3672
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   8162
   OleObjectBlob   =   "frmInputBoxPlus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInputBoxPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mclsResizer As clsFormResizer
Public m_ReturnValue As Variant
Public m_ButtonClicked As String
Private m_PasswordMode As Boolean
Private m_RegExMode As Boolean
Private m_RefEditMode As Boolean

Sub Initialise(Prompt As String, Title As String, Default As String, OKText As String, CancelText As String, _
        TextBoxWidth As Double, TextBoxHeight As Double, PasswordMode As Boolean, RegExMode As Boolean, MiddleText As String, RefEditMode As Boolean)

          Const LeftMargin = 5
          Const TopMargin = 5
          Const vGap = 10
          Dim Accelerator As String
          Dim ButtonHeight As Double
          Dim ButtonWidth As Double
          Dim c As control
          Dim ControlToUse As control
          Dim Have3Buttons As Boolean

1         On Error GoTo ErrHandler
          
2         For Each c In Me.Controls
3             c.Font.Name = "Segoe UI"
4             c.Font.Size = 9
5         Next c

6         If MiddleText = vbNullString Then
7             Have3Buttons = False
8             but3.Visible = False
9             butOK.Default = True
10            butOK.Cancel = False
11            butCancel.Cancel = True
12            butCancel.Default = False
13        Else
14            Have3Buttons = True
15            but3.Default = True
16            but3.Cancel = False
17            butOK.Default = False
18            butOK.Cancel = False
19            butCancel.Cancel = True
20            butCancel.Default = False
21        End If

22        With Label1
23            .caption = Prompt
24            .Width = 1000
25            .AutoSize = False
26            .AutoSize = True
27            .Left = LeftMargin
28            .Top = TopMargin
29        End With

          'Pitfalls of RefEdit controls:
          'http://peltiertech.com/using-refedit-controls-in-excel-dialogs/

30        If RefEditMode Then
31            Set ControlToUse = RefEdit1
32            TextBox1.Visible = False
33        Else
34            Set ControlToUse = TextBox1
35            RefEdit1.Visible = False
36        End If

37        With ControlToUse
38            .Value = Default
39            .Width = TextBoxWidth
40            .Height = TextBoxHeight
41            .Top = Label1.Top + Label1.Height + vGap
42            .Left = LeftMargin
43        End With

44        m_PasswordMode = PasswordMode
45        m_RegExMode = RegExMode
46        m_RefEditMode = RefEditMode
47        If PasswordMode Then
48            TextBox1.PasswordChar = Chr$(149)
49            LabelShow.Visible = True
50        Else
51            LabelShow.Visible = False
52        End If

53        Me.caption = Title

54        With butOK
55            .caption = ProcessAmpersands(OKText, Accelerator)
56            If Accelerator <> vbNullString Then .Accelerator = Accelerator
57            .Width = 500
58            .AutoSize = False
59            .AutoSize = True
60        End With

61        With butCancel
62            .caption = ProcessAmpersands(CancelText, Accelerator)
63            If Accelerator <> vbNullString Then .Accelerator = Accelerator
64            .Width = 500
65            .AutoSize = False
66            .AutoSize = True
67        End With

68        If Have3Buttons Then
69            With but3
70                .caption = ProcessAmpersands(MiddleText, Accelerator)
71                If Accelerator <> vbNullString Then .Accelerator = Accelerator
72                .Width = 500
73                .AutoSize = False
74                .AutoSize = True
75            End With
76        End If

77        If PasswordMode Then
78            With LabelShow
79                .Left = LeftMargin
80                .Top = TextBox1.Top + TextBox1.Height
81                .Tag = "T"
82            End With
83        End If

84        If Have3Buttons Then
85            ButtonWidth = SafeMax(SafeMax(78, but3.Width), SafeMax(butOK.Width, butCancel.Width))
86            ButtonHeight = SafeMax(SafeMax(36, but3.Height), SafeMax(butOK.Height, butCancel.Height))
87        Else
88            ButtonWidth = SafeMax(78, SafeMax(butOK.Width, butCancel.Width))
89            ButtonHeight = SafeMax(36, SafeMax(butOK.Height, butCancel.Height))
90        End If

91        With butOK
92            .Width = ButtonWidth
93            .Height = ButtonHeight
94            .Left = LeftMargin
95            If PasswordMode Then
96                .Top = LabelShow.Top + LabelShow.Height + vGap
97            Else
98                .Top = ControlToUse.Top + ControlToUse.Height + vGap
99            End If
100       End With

101       If Have3Buttons Then
102           With but3
103               .Width = ButtonWidth
104               .Height = ButtonHeight
105               .Left = butOK.Left + butOK.Width + vGap
106               .Top = butOK.Top
107           End With
108       End If

109       With butCancel
110           .Width = ButtonWidth
111           .Height = ButtonHeight
112           If Have3Buttons Then
113               .Left = but3.Left + but3.Width + vGap
114           Else
115               .Left = butOK.Left + butOK.Width + vGap
116           End If
117           .Top = butOK.Top
118       End With

119       Me.Height = butOK.Top + butOK.Height + (Me.Height - Me.InsideHeight) + 11.35
          'butCancel.Width + 5 in line below is to leave space for the resizing handle

120       Me.Width = SafeMax(SafeMax(butCancel.Left + butCancel.Width + 5, Label1.Left + Label1.Width), ControlToUse.Left + ControlToUse.Width) + (Me.Width - Me.InsideWidth) + 11.35

121       If RefEditMode Then
122           RefEdit1.SetFocus
123       Else
124           With TextBox1
125               .SetFocus
126               .SelStart = 0
127               .SelLength = Len(TextBox1.Value)
128           End With
129       End If

          'Set up resizer...
130       If RefEditMode Then
131           RefEdit1.Tag = "W"
132           Label1.Tag = "W"
133       Else
134           butOK.Tag = "T"
135           butCancel.Tag = "T"
136           TextBox1.Tag = "WH"
137           Label1.Tag = "W"
138       End If

139       CreateFormResizer mclsResizer
140       mclsResizer.Initialise Me, Me.Height, Me.Width, Me.BackColor

141       Exit Sub
ErrHandler:
142       SomethingWentWrong "#Initialise (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butCancel_Click()
1         On Error GoTo ErrHandler
2         m_ReturnValue = False
3         m_ButtonClicked = butCancel.caption
4         HideForm Me
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butCancel_Click (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_Click()
1         On Error GoTo ErrHandler
2         If m_RefEditMode Then
3             m_ReturnValue = RefEdit1.Value
4         Else
5             m_ReturnValue = TextBox1.Value
6         End If
7         m_ButtonClicked = butOK.caption
8         HideForm Me
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#butOK_Click (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub but3_Click()
1         On Error GoTo ErrHandler
2         If m_RefEditMode Then
3             m_ReturnValue = RefEdit1.Value
4         Else
5             m_ReturnValue = TextBox1.Value
6         End If
7         m_ButtonClicked = but3.caption
8         HideForm Me
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#but3_Click (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butOK
3         UnHighlightFormControl butCancel
4         UnHighlightFormControl but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmInputBoxPlus.butOK_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmInputBoxPlus.butCancel_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub but3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl but3
3         UnHighlightFormControl butOK
4         UnHighlightFormControl butCancel
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmInputBoxPlus.but3_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub LabelShow_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         If m_PasswordMode Then
3             TextBox1.PasswordChar = vbNullString
4         End If
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#LabelShow_MouseDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub LabelShow_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         If m_PasswordMode Then
3             TextBox1.PasswordChar = Chr$(149)
4         End If
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#LabelShow_MouseUp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
          Dim Res
1         On Error GoTo ErrHandler
2         If m_RegExMode Then
3             Cancel.Value = True
4             Res = ShowRegularExpressionDialog(TextBox1.Value, "Formula", , Me, "Search formulas", "Find formulas where", True, "SearchFormulas")
5             If Res <> "#User Cancel!" Then
6                 TextBox1.Value = Res
7             End If
8         End If
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#TextBox1_DblClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but3

5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

1         On Error GoTo ErrHandler

2         If CloseMode = vbFormControlMenu Then
3             butCancel_Click
4         End If

5         Exit Sub

ErrHandler:
6         Throw "#frmInputBoxPlus.UserForm_QueryClose(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

