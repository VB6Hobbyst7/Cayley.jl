VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionButton 
   Caption         =   "Select"
   ClientHeight    =   4770
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   6720
   OleObjectBlob   =   "frmOptionButton.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private EventHandlers() As New clsOptionButton
Private Const m_FontName = "Segoe UI"
Private Const m_MinFrameHeight = 20
Private Const m_FontSize = 9
Private Const m_MaxButtonsInColumn = 8
Private Const m_vSpacer = 2        'extra vertical space between the option buttons
Private Const m_hSpacer = 5        'extra horizontal space between the option buttons
Private m_NumOptions As Long
Public m_NumGroups As Long
Private m_HelpMethodName As String
Public ChosenIndices As Variant
Public ChosenValues As Variant
Public ButtonClicked As Variant
Private AcceleratorsUsed() As Long

Sub Initialise(ByVal TheChoices, Optional Title As String, Optional CurrentChoice As Variant, Optional TopText As Variant, _
        Optional CheckBoxText As String, Optional CheckBoxValue As Boolean, Optional HelpMethodName As String, _
        Optional Caption1 As String = "OK", Optional Caption2 As String, Optional Caption3 As String = "Cancel")

1         On Error GoTo ErrHandler

          Dim Accelerator As String
          Dim Anchor As control
          Dim butCount As Long
          Dim ChooseVector
          Dim ColumnWidth As Double
          Dim FoundMatch As Boolean
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim MaxHeight As Double
          Dim NR As Long
          Dim opBut As control
          Dim OpButAtTopOfCurrentCol As control
          Dim PreviousColumnWidth As Double
          Dim prevOpBut As control
          Dim startNextColumn As Boolean
          Dim TheFrame As control
          Dim TheLabel As control
          Dim TheseChoices
          Dim WidthOfe As Double

2         Me.ChosenIndices = Empty: Me.ChosenValues = Empty
3         If Title <> vbNullString Then Me.caption = Title
4         Force2DArrayR TheChoices, NR

5         m_NumGroups = sNCols(TheChoices)
6         ChooseVector = sReshape(False, NR, 1)
7         m_NumOptions = 0
8         For i = 1 To NR
9             For j = 1 To m_NumGroups
10                If Not IsError(TheChoices(i, j)) Then m_NumOptions = m_NumOptions + 1
11            Next
12        Next

13        ReDim AcceleratorsUsed(0 To 255)
14        Me.Font.Name = m_FontName
15        WidthOfe = sStringWidth("e", m_FontName, m_FontSize)(1, 1)

16        ReDim EventHandlers(1 To m_NumOptions)

17        For j = 1 To m_NumGroups

18            If j = 1 Then
19                Set TheLabel = lblTopText
20            Else
21                Set TheLabel = Me.Controls.Add("Forms.Label.1")
22            End If
23            If TopText(1, j) <> vbNullString Then
24                With TheLabel
25                    .Visible = True
26                    .Font.Name = m_FontName
27                    .Font.Size = m_FontSize
28                    .Visible = True
29                    .Width = SafeMax(Frame1.Width, 50 * WidthOfe)    'TODO this is poor: Frame1.width not correctly set yet
30                    .caption = TopText(1, j) + " |" 'AutoSize is flakey. Size for slighly more text.
31                    .AutoSize = False
32                    .AutoSize = True
33                    .AutoSize = False
34                    .caption = TopText(1, j)
35                    If j = 1 Then
36                        .Left = 6
37                        .Top = 6
38                    Else
39                        With Me.Controls("Frame" & j - 1)
40                            TheLabel.Top = .Top + .Height + 6
41                            TheLabel.Left = 6
42                        End With
43                    End If
44                End With
45            Else
46                TheLabel.Visible = False
47            End If

48            If m_NumGroups = 1 Then
49                TheseChoices = TheChoices
50            Else
51                TheseChoices = sSubArray(TheChoices, 1, j, , 1)
52                For k = 1 To NR
53                    ChooseVector(k, 1) = Not (IsError(TheseChoices(k, 1)))
54                Next k
55                TheseChoices = sMChoose(TheseChoices, ChooseVector)
56            End If
57            If Not (IsNumeric(sMatch(CurrentChoice(1, j), TheseChoices))) Then
58                CurrentChoice(1, j) = TheseChoices(1, 1)
59            End If

60            FoundMatch = False
61            ColumnWidth = 0
62            MaxHeight = 0
63            If j = 1 Then
64                Set TheFrame = Frame1
65            Else
66                Set TheFrame = Me.Controls.Add("Forms.Frame.1", "Frame" & j)
67            End If

              Dim NL As Long
68            For i = 1 To sNRows(TheseChoices)
69                butCount = butCount + 1

70                Set opBut = TheFrame.Controls.Add("Forms.OptionButton.1", "OptionButton" & j & "_" & i)
71                Set EventHandlers(butCount).butEvents = opBut
72                EventHandlers(butCount).Tag = butCount

73                With opBut
74                    .TabStop = True
75                    .TabIndex = butCount
76                    .Font.Name = m_FontName
77                    .Font.Size = m_FontSize
78                    .Visible = True
79                    .Accelerator = ChooseAccelerator(CStr(TheseChoices(i, 1)))
80                    .caption = ProcessAmpersands(CStr(TheseChoices(i, 1)), vbNullString)
81                    .Tag = CStr(TheseChoices(i, 1))    'So that we can reliably populate ChosenValues at exit
82                    If TheseChoices(i, 1) = CurrentChoice(1, j) Then
83                        If Not FoundMatch Then
84                            .Value = True
85                            .SetFocus
86                            FoundMatch = True
87                        End If
88                    End If
89                    If InStr(TheseChoices(i, 1), vbLf) = 0 Then
90                        NL = 1
91                        .Width = sStringWidth(TheseChoices(i, 1), .Font.Name, .Font.Size)(1, 1) + 25
92                    Else
93                        .Width = sMaxOfArray(sStringWidth(sTokeniseString(CStr(TheseChoices(i, 1)), vbLf), .Font.Name, .Font.Size)) + 25
94                        NL = sNRows(sTokeniseString(CStr(TheseChoices(i, 1)), vbLf))
95                    End If

96                    ColumnWidth = SafeMax(ColumnWidth, .Width)
97                    .Height = 19 * NL 'previously used 18 but occasionally caption was appearing in too small a font
98                    If startNextColumn Then
99                        PreviousColumnWidth = ColumnWidth
100                       ColumnWidth = .Width
101                   End If
102                   If i > 1 Then
103                       Set prevOpBut = Me.Controls("OptionButton" & j & "_" & i - 1)
104                       If startNextColumn Then
105                           startNextColumn = False
106                           .Top = OpButAtTopOfCurrentCol.Top
107                           .Left = OpButAtTopOfCurrentCol.Left + PreviousColumnWidth + m_hSpacer
108                           Set OpButAtTopOfCurrentCol = opBut
109                       Else
110                           .Left = prevOpBut.Left
111                           .Top = prevOpBut.Top + prevOpBut.Height + m_vSpacer
112                       End If
113                   Else
114                       Set OpButAtTopOfCurrentCol = opBut
115                       .Top = m_vSpacer * 0.5
116                       .Left = m_hSpacer
117                   End If
118                   MaxHeight = SafeMax(MaxHeight, .Top + .Height)
119                   If (i Mod m_MaxButtonsInColumn) = 0 Then
120                       startNextColumn = True
121                   End If
122               End With
123           Next i
124           TheFrame.Height = SafeMax(m_MinFrameHeight, MaxHeight + 2 * m_vSpacer)
125           TheFrame.Width = opBut.Left + ColumnWidth + m_hSpacer
126           If j = 1 Then
127               If TheLabel.Visible Then
128                   TheFrame.Top = TheLabel.Top + TheLabel.Height + m_vSpacer
129                   TheFrame.Left = TheLabel.Left
130               End If
131           Else
132               If TheLabel.Visible Then
133                   Set Anchor = TheLabel
134               Else
135                   Set Anchor = Me.Controls("Frame" & (j - 1))
136               End If
137               TheFrame.Top = Anchor.Top + Anchor.Height + m_vSpacer
138               TheFrame.Left = Anchor.Left
139           End If
140       Next j
141       Frame1.ActiveControl.SetFocus

142       With Me
143           .butOK.Top = TheFrame.Top + TheFrame.Height + 5
144           .butCancel.Top = .butOK.Top
145           .butOK.Font.Name = m_FontName
146           .butOK.Font.Size = m_FontSize
147           .butCancel.Font.Name = m_FontName
148           .butCancel.Font.Size = m_FontSize
149           .but2.Font.Name = m_FontName
150           .but2.Font.Size = m_FontSize
151           .butOK.caption = ProcessAmpersands(Caption1, Accelerator)
152           If Accelerator <> vbNullString Then butOK.Accelerator = Accelerator
153           .butCancel.caption = ProcessAmpersands(Caption3, Accelerator)
154           If Accelerator <> vbNullString Then butCancel.Accelerator = Accelerator
155           If Caption2 <> vbNullString Then
156               but2.Visible = True
157               but2.caption = ProcessAmpersands(Caption2, Accelerator)
158               If Accelerator <> vbNullString Then but2.Accelerator = Accelerator
159               but2.Left = butOK.Left + butOK.Width + 6
160               butCancel.Left = but2.Left + but2.Width + 6
161               but2.Top = butOK.Top
162           Else
163               but2.Visible = False
164               butCancel.Left = butOK.Left + butOK.Width + 6
165           End If
166           butOK.TabStop = False
167           but2.TabStop = False
168           butCancel.TabStop = False
169           CheckBox1.TabStop = False
170           .Height = butOK.Top + butOK.Height + m_vSpacer * 2 + (Me.Height - Me.InsideHeight)
171       End With

          Dim FrameWidth As Double

172       For j = 1 To m_NumGroups
173           FrameWidth = SafeMax(FrameWidth, Me.Controls("Frame" & j).Width)
174       Next
175       FrameWidth = SafeMax(FrameWidth, butCancel.Left + butCancel.Width - butOK.Left)

176       For j = 1 To m_NumGroups
177           Me.Controls("Frame" & j).Width = FrameWidth
178       Next j
          Dim borderWidth As Double
179       borderWidth = (Me.Width - Me.InsideWidth)

180       Me.Width = SafeMax(Frame1.Left + FrameWidth + m_hSpacer + borderWidth, butCancel.Left + butCancel.Width + m_hSpacer + borderWidth)

181       If CheckBoxText <> vbNullString Then
182           With CheckBox1
183               .Font.Name = m_FontName
184               .Font.Size = m_FontSize
185               .caption = ProcessAmpersands(CheckBoxText, vbNullString)
186               .Accelerator = ChooseAccelerator(CheckBoxText)
187               .Value = CheckBoxValue
188               .Width = Frame1.Width - m_hSpacer    'This is poor - Frame1 width has not yet been correctly set
189               .AutoSize = True
190               .AutoSize = False
191               .Top = TheFrame.Top + TheFrame.Height + 5
192               .Left = Frame1.Left + m_hSpacer
193               Me.Height = Me.Height + .Height + 10
194               butOK.Top = butOK.Top + .Height + 10
195               butCancel.Top = butOK.Top
196               but2.Top = butOK.Top
197           End With
198       Else
199           CheckBox1.Visible = False
200       End If

201       m_HelpMethodName = HelpMethodName
202       If Len(HelpMethodName) > 0 Then
203           With lblHelp
204               .Font.Name = m_FontName
205               .Font.Size = m_FontSize
206               .Width = 100
207               .Height = 100
208               .AutoSize = False
209               .AutoSize = True
210               .Visible = True
211               .Left = butOK.Left
212               .Top = butOK.Top + butOK.Height + 10
213               Me.Height = Me.Height + .Height + 10
214           End With
215       Else
216           lblHelp.Visible = False
217       End If

218       Exit Sub
ErrHandler:
219       Throw "#Initialise(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CapitaliseAndStripAccents
' Author    : Philip Swannell
' Date      : 16-Nov-2015
' Purpose   : Morphs a > A, é to E etc. We need this since if the underlined character
'             for one of the options has an accent then we need the code to respond to
'             the unaccented equivalent character.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CapitaliseAndStripAccents(Character As String)
1         On Error GoTo ErrHandler
2         If Len(Character) = 0 Then
3             CapitaliseAndStripAccents = vbNullString
4         Else
5             Select Case Asc(Left$(Character, 1))
                  Case Is < 138
6                     CapitaliseAndStripAccents = UCase$(Left$(Character, 1))
7                 Case 192, 193, 194, 195, 196, 197, 224, 225, 226, 227, 228, 229
8                     CapitaliseAndStripAccents = "A"
9                 Case 199, 231
10                    CapitaliseAndStripAccents = "C"
11                Case 208, 240
12                    CapitaliseAndStripAccents = "D"
13                Case 200, 201, 202, 203, 232, 233, 234, 235
14                    CapitaliseAndStripAccents = "E"
15                Case 204, 205, 206, 207, 236, 237, 238, 239
16                    CapitaliseAndStripAccents = "I"
17                Case 209, 241
18                    CapitaliseAndStripAccents = "N"
19                Case 210, 211, 212, 213, 214, 242, 243, 244, 245, 246
20                    CapitaliseAndStripAccents = "O"
21                Case 138, 154
22                    CapitaliseAndStripAccents = "S"
23                Case 217, 218, 219, 220, 249, 250, 251, 252
24                    CapitaliseAndStripAccents = "U"
25                Case 159, 221, 253, 255
26                    CapitaliseAndStripAccents = "Y"
27                Case 142, 158
28                    CapitaliseAndStripAccents = "Z"
29                Case Else
30                    CapitaliseAndStripAccents = UCase$(Left$(Character, 1))
31            End Select
32        End If
33        Exit Function
ErrHandler:
34        Throw "#CapitaliseAndStripAccents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ChooseAccelerator
' Author    : Philip Swannell
' Date      : 14-Nov-2015
' Purpose   : Choose an accelerator key for the option, we take the first character that
'             has not been used so far. If all have been used as accelerators for previous
'             option buttons then we take the character that has been least used.
'             BUT if passed in caption has & character, we strip that & from the caption
'             displayed and use the following character as the accelerator.
'TODO: Handle accented characters gracefully...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ChooseAccelerator(ByVal caption As String) As String
          Dim charForMinUsage As String
          Dim DollarMatch As Long
          Dim i As Long
          Dim minUsage
          Dim ThisChar As String
          Dim ThisUsage As Long
1         On Error GoTo ErrHandler

2         caption = Replace(caption, "&&", vbNullString)        '&& means we want a single $ to appear on the screen, and it's not an indicator that the following character should be the accelerator

3         DollarMatch = InStr(caption, "&")
4         If DollarMatch > 0 Then
5             If DollarMatch < Len(caption) Then
6                 ChooseAccelerator = Mid$(caption, DollarMatch + 1, 1)
7                 AcceleratorsUsed(Asc(CapitaliseAndStripAccents(ChooseAccelerator))) = AcceleratorsUsed(Asc(CapitaliseAndStripAccents(ChooseAccelerator))) + 1
8                 Exit Function
9             End If
10        End If

11        minUsage = 100

12        For i = 1 To Len(caption)
13            ThisChar = Mid$(caption, i, 1)
14            If ThisChar <> " " And ThisChar <> "." And ThisChar <> "," Then
15                ThisUsage = AcceleratorsUsed(Asc(CapitaliseAndStripAccents(ThisChar)))
16                If ThisUsage = 0 Then
17                    AcceleratorsUsed(Asc(CapitaliseAndStripAccents(ThisChar))) = 1
18                    ChooseAccelerator = ThisChar
19                    Exit Function
20                End If
21                If ThisUsage < minUsage Then
22                    minUsage = ThisUsage
23                    charForMinUsage = ThisChar
24                End If
25            End If
26        Next i
          Dim cleanedCharacter As String

27        If Len(charForMinUsage) > 0 Then
28            cleanedCharacter = CapitaliseAndStripAccents(charForMinUsage)
29            AcceleratorsUsed(Asc(cleanedCharacter)) = AcceleratorsUsed(Asc(cleanedCharacter)) + 1
30        End If

31        ChooseAccelerator = charForMinUsage
32        Exit Function
ErrHandler:
33        Throw "#ChooseAccelerator (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub but2_Click()
1         On Error GoTo ErrHandler
2         Clicked_OK
3         ButtonClicked = but2.caption
4         HideForm Me
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#but2_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butCancel_Click()
1         On Error GoTo ErrHandler
2         ButtonClicked = butCancel.caption
3         HideForm Me
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butCancel_Click(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_Click()
1         On Error GoTo ErrHandler
2         ButtonClicked = butOK.caption
3         Clicked_OK
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#butOk_Click(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

'Public so that can be called from clsOptionButton
Public Sub Clicked_OK()
1         On Error GoTo ErrHandler
          Dim actrl As control
          Dim i As Long
          Dim TmpI
          Dim TmpV

2         TmpI = sReshape(0, 1, m_NumGroups)
3         TmpV = sReshape(vbNullString, 1, m_NumGroups)

4         For i = 1 To m_NumGroups
5             For Each actrl In Me.Controls("Frame" & i).Controls
6                 If actrl.Value Then
7                     TmpI(1, i) = CInt(sStringBetweenStrings(actrl.Name, "_"))
8                     TmpV(1, i) = actrl.Tag
9                     Exit For
10                End If
11            Next
12        Next i
13        Me.ChosenIndices = TmpI
14        Me.ChosenValues = TmpV
15        HideForm Me
16        Exit Sub
ErrHandler:
17        Throw "#Clicked_OK(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : KeyPressResponse
' Author    : Philip Swannell
' Date      : 15-Nov-2015
' Purpose   : The option buttons have accelerator keys. i.e. the button gets focus if the user hits Alt character
'             but this routine makes it the case that simply hitting the accelerator key (without Alt) is equivalent to
'             selecting that option and clicking OK.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub KeyPressResponse(KeyAscii As MSForms.ReturnInteger)

          Dim c As control
          Dim i As Long
          Dim N As Long
          Dim NumberChosen As Long

1         On Error GoTo ErrHandler

2         If m_NumGroups = 1 Then    'This "quick response to key press" only makes sense if the user has exactly one choice to make, i.e. there is just one group and no check box
3             If CheckBox1.Visible = False Then

4                 For i = 1 To Frame1.Controls.Count
5                     Set c = Me.Controls("OptionButton1_" & CStr(i))
6                     If CapitaliseAndStripAccents(c.Accelerator) = CapitaliseAndStripAccents(Chr$(KeyAscii)) Then
7                         N = N + 1
8                         NumberChosen = i
9                     End If
10                Next i

11                If N = 0 Then
12                    Exit Sub
13                ElseIf N = 1 Then
14                    Me.ChosenIndices = sReshape(NumberChosen, 1, 1)    'Needs to be 2-dimensional array
15                    HideForm Me
16                Else
                      'More than one option has the same accelerator so carrousel between them...
                      Dim First As Long
17                    First = CLng(Replace(Frame1.ActiveControl.Name, "OptionButton1_", vbNullString))
18                    For i = 1 To m_NumOptions
19                        Set c = Me.Controls("OptionButton1_" + CStr((First + i - 1) Mod m_NumOptions + 1))
20                        If CapitaliseAndStripAccents(c.Accelerator) = CapitaliseAndStripAccents(Chr$(KeyAscii)) Then
21                            c.SetFocus
22                            c.Value = True
23                            Exit For
24                        End If
25                    Next i
26                End If
27            End If
28        End If

29        Exit Sub
ErrHandler:
30        SomethingWentWrong "#KeyPressResponse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'NB Rubberduck incorrecty states that this method is not used. It is, from clsOptionButton
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : KeyDownResponse
' Author    : Philip Swannell
' Date      : 17-Nov-2015
' Purpose   : Make the arrow keys traverse the option buttons in an intuitive way, paying
'             attention to their position on the screen.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub KeyDownResponse(ByVal KeyCode As MSForms.ReturnInteger, Shift As Integer)
          Dim nextOptionButtonNumber As Long
          Dim NumInGroup As Long
          Dim ThisGroupNumber As Long
          Dim thisOptionButtonNumber As Long
1         On Error GoTo ErrHandler

          ' PGS 1 Nov 2016
          ' It appears that there are differences in how option buttons on forms repond to the user hitting an arrow key
          ' as between Office 2016 and earlier versions of Office. If there are a number of option buttons inside a group box
          ' and the user hits an arrow key.

          ' In Office 2013 another option button takes focus (text of option button has a dotted border)
          ' but the currently-selected remains the one with the black dot (Value property is True)
          ' In Office 2016 another option button takes focus (text of option button has a dotted border)
          ' and also becomes the one with the black dot (value property True)

          ' The Office 2016 behaviour is "better", perhaps a bug fix, and makes this method redundant, in fact this method
          ' would "fight" with the in-built reponse to key press such that pressing (say) the down key selects the next-but-one
          ' option button, which is bad.

2         If Val(Application.Version) >= 16 Then Exit Sub

3         thisOptionButtonNumber = CLng(sStringBetweenStrings(Me.ActiveControl.ActiveControl.Name, "_"))
4         ThisGroupNumber = CLng(sStringBetweenStrings(Me.ActiveControl.ActiveControl.Name, "OptionButton", "_"))
5         NumInGroup = Me.Controls("Frame" & ThisGroupNumber).Controls.Count

6         If KeyCode = 40 Or (KeyCode = 9 And Shift = 0) Then        'down arrow or Tab - traverse like Chinese reading
7             If thisOptionButtonNumber = NumInGroup Then
8                 nextOptionButtonNumber = 1
9             Else
10                nextOptionButtonNumber = thisOptionButtonNumber + 1
11            End If
12        ElseIf KeyCode = 38 Or (KeyCode = 9 And Shift = 1) Then        'up arrow or Shift Tab - traverse like backwards Chinese reading
13            If thisOptionButtonNumber = 1 Then
14                nextOptionButtonNumber = NumInGroup
15            Else
16                nextOptionButtonNumber = thisOptionButtonNumber - 1
17            End If
18        ElseIf KeyCode = 39 Then        'right arrow - traverse like Western reading
19            If thisOptionButtonNumber <= NumInGroup - m_MaxButtonsInColumn Then
20                nextOptionButtonNumber = thisOptionButtonNumber + m_MaxButtonsInColumn
21            Else
22                nextOptionButtonNumber = thisOptionButtonNumber Mod m_MaxButtonsInColumn + 1
23                If nextOptionButtonNumber > NumInGroup Then
24                    nextOptionButtonNumber = 1
25                End If
26            End If
27        ElseIf KeyCode = 37 Then        'left arrow - traverse like backwards Western reading.
28            If thisOptionButtonNumber > m_MaxButtonsInColumn Then
29                nextOptionButtonNumber = thisOptionButtonNumber - m_MaxButtonsInColumn
30            Else
31                For nextOptionButtonNumber = NumInGroup To NumInGroup - m_MaxButtonsInColumn + 1 Step -1
32                    If nextOptionButtonNumber Mod m_MaxButtonsInColumn = (thisOptionButtonNumber + 1) Mod m_MaxButtonsInColumn Then Exit For
33                Next
34                If nextOptionButtonNumber < 1 Then
35                    nextOptionButtonNumber = 1
36                End If
37            End If
38        Else
39            Exit Sub
40        End If
41        With Me.Controls("OptionButton" & ThisGroupNumber & "_" + CStr(nextOptionButtonNumber))
42            .SetFocus
43            .Value = True
44        End With
45        Exit Sub
ErrHandler:
46        SomethingWentWrong "#KeyDownResponse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub lblHelp_Click()
1         On Error GoTo ErrHandler
2         Application.Run m_HelpMethodName
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#lblHelp_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but2
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butOK
3         UnHighlightFormControl butCancel
4         UnHighlightFormControl but2
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butOK_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub butCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl butCancel
3         UnHighlightFormControl butOK
4         UnHighlightFormControl but2
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#butCancel_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
1         On Error GoTo ErrHandler
2         If (CloseMode = vbFormControlMenu) Then
3             Cancel = CLng(True)
4             HideForm Me
5         End If
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#UserForm_QueryClose(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub butOK_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         KeyPressResponse KeyAscii
End Sub

Private Sub butCancel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         KeyPressResponse KeyAscii
End Sub

Private Sub but2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         KeyPressResponse KeyAscii
End Sub

Private Sub but2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl but2
3         UnHighlightFormControl butOK
4         UnHighlightFormControl butCancel
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#but2_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, Me.caption
End Sub

