VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgBoxPlus 
   Caption         =   "Microsoft Excel"
   ClientHeight    =   6255
   ClientLeft      =   42
   ClientTop       =   392
   ClientWidth     =   9394.001
   OleObjectBlob   =   "frmMsgBoxPlus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMsgBoxPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mclsResizer As clsFormResizer
Private m_but1Result As VbMsgBoxResult
Private m_but2Result As VbMsgBoxResult
Private m_but3Result As VbMsgBoxResult
Public m_ReturnValue As VbMsgBoxResult
Public m_ReturnCaption As String
Private m_ButtonsSimple As VbMsgBoxStyle
Private m_numButsVisible As Long
Private Const m_FontName = "Segoe UI"
Private Const m_FontSize = 9
Private Const m_GapAboveLabel As Long = 20
Private Const m_GapAboveButtons = 15
Private Const m_GapOnLeft = 15
Private Const m_GapOnRight = 30
Private Const m_GapBetweenButtons As Long = 5
Private Const m_GapAboveCheckBox = 5
Private Const m_MaxLabelHeight As Long = 600
Private Const m_CheckBoxNudge As Double = 8
Private Const ButtonsError As String = "Buttons must be sum of A, B and C where A is either vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo or vbRetryCancel;B either 0, vbCritical, vbQuestion, vbExclamation or vbInformation; C is either vbDefaultButton1, vbDefaultButton2 or vbDefaultButton3"
Private m_AcceleratorsUsed As String
Private m_HideNow As Boolean
Private m_SelfDestructButton As VbMsgBoxResult
Private m_SelfDestructMode As Boolean
Private m_SecondsToSelfDestruct As Long

'PGS 24/11/2015 used Microsoft Office 2010 Code Compatibility Inspector to make changes for 64-bit compatibility
Private Declare PtrSafe Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long

'vbOKOnly            0
'vbOKCancel          1
'vbAbortRetryIgnore  2
'vbYesNoCancel       3
'vbYesNo             4
'vbRetryCancel       5
'vbCritical          16
'vbQuestion          32
'vbExclamation       48
'vbInformation       64
'vbDefaultButton1    0
'vbDefaultButton2    256
'vbDefaultButton3    512
'vbDefaultButton4    768

Sub Initialise(Optional ByVal Prompt As String, _
          Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
          Optional ByVal Title As String = "Microsoft Excel", _
          Optional Caption1 As String = vbNullString, _
          Optional Caption2 As String = vbNullString, _
          Optional Caption3 As String = vbNullString, _
          Optional ButtonWidth As Double = 70, _
          Optional TextWidth As Long = 300, _
          Optional ByVal CheckBoxCaption As String = vbNullString, _
          Optional ByRef CheckBoxValue As Boolean, _
          Optional ByVal SecondsToSelfDestruct As Long, _
          Optional SelfDestructButton As VbMsgBoxResult)

          Dim Ctrl As control
          Dim DefaultButtonNumber As Long
          Dim i As Long
          Dim imgToShow As control
          Dim NumDefaultButtons As Long
          Dim numVis As Long

1         On Error GoTo ErrHandler
          'Parse Buttons argument. Need to test for components of Buttons from largest to smallest - see list at top of module.
2         If Buttons > (vbDefaultButton3 + vbInformation + vbRetryCancel) Then Throw ButtonsError
3         m_ButtonsSimple = Buttons

4         If (m_ButtonsSimple And vbDefaultButton3) <> 0 And (m_ButtonsSimple >= vbDefaultButton3) Then
5             m_ButtonsSimple = m_ButtonsSimple - vbDefaultButton3
6             DefaultButtonNumber = 3
7             NumDefaultButtons = NumDefaultButtons + 1
8         End If
9         If (m_ButtonsSimple And vbDefaultButton2) <> 0 And (m_ButtonsSimple >= vbDefaultButton2) Then
10            m_ButtonsSimple = m_ButtonsSimple - vbDefaultButton2
11            DefaultButtonNumber = 2
12            NumDefaultButtons = NumDefaultButtons + 1
13        End If
14        If NumDefaultButtons > 1 Then
15            Throw ButtonsError
16        ElseIf NumDefaultButtons = 0 Then
17            DefaultButtonNumber = 1
18            NumDefaultButtons = 1
19        End If

          'If we have images at the correct resolution, then use them, otherwise stretch the highest resolution lower than what we want.
          'Bitmaps are at PGS OneDrive at \Excel Sheets\Bitmaps. They can be extracted from c:\Windows\System32\imageres.dll
          'using BeCyIconGrabber.exe available at https://download.cnet.com/BeCyIconGrabber/3000-2192_4-10768921.html
          Dim DPI As Long
          Dim ImageSuffix As String
          Dim SizeMode As Long
20        For Each Ctrl In Me.Controls
21            If Left$(Ctrl.Name, 3) = "img" Then Ctrl.Visible = False
22        Next
23        DPI = ScreenDPI(True)
24        Select Case DPI
              Case 96, 120, 144, 192, 216
25                ImageSuffix = CStr(DPI / 3)
26                SizeMode = 0 'frmPictureSizeModeClip
27            Case Is > 216
28                ImageSuffix = CStr(216 / 3)
29                SizeMode = 1 'frmPicureSizeModeStretch
30            Case Is > 192
31                ImageSuffix = CStr(192 / 3)
32                SizeMode = 1 'frmPicureSizeModeStretch
33            Case Is > 144
34                ImageSuffix = CStr(144 / 3)
35                SizeMode = 1 'frmPicureSizeModeStretch
36            Case Is > 120
37                ImageSuffix = CStr(120 / 3)
38                SizeMode = 1 'frmPicureSizeModeStretch
39            Case Else
40                ImageSuffix = CStr(96 / 3)
41                SizeMode = 1 'frmPicureSizeModeStretch
42        End Select

43        If (m_ButtonsSimple And vbInformation) <> 0 And (m_ButtonsSimple >= vbInformation) Then
44            m_ButtonsSimple = m_ButtonsSimple - vbInformation
45            numVis = numVis + 1
46            Set imgToShow = Me.Controls("imgInformation" & ImageSuffix)
47        End If
48        If (m_ButtonsSimple And vbExclamation) <> 0 And (m_ButtonsSimple >= vbExclamation) Then
49            m_ButtonsSimple = m_ButtonsSimple - vbExclamation
50            numVis = numVis + 1
51            Set imgToShow = Me.Controls("imgExclamation" & ImageSuffix)
52        End If
53        If (m_ButtonsSimple And vbQuestion) <> 0 And (m_ButtonsSimple >= vbQuestion) Then
54            m_ButtonsSimple = m_ButtonsSimple - vbQuestion
55            numVis = numVis + 1
56            Set imgToShow = Me.Controls("imgQuestion" & ImageSuffix)
57        End If
58        If (m_ButtonsSimple And vbCritical) <> 0 And (m_ButtonsSimple >= vbCritical) Then
59            m_ButtonsSimple = m_ButtonsSimple - vbCritical
60            numVis = numVis + 1
61            Set imgToShow = Me.Controls("imgCritical" & ImageSuffix)
62        End If
63        If numVis > 1 Then
64            Throw ButtonsError
65        ElseIf numVis = 1 Then
66            With imgToShow
67                .Visible = True
68                .Width = 24
69                .Height = 24
70                .PictureSizeMode = SizeMode
71            End With
72        End If

73        Select Case m_ButtonsSimple
              Case vbOKOnly
74                m_but1Result = vbOK
75                m_numButsVisible = 1
76            Case vbOKCancel
77                m_but1Result = vbOK
78                m_but2Result = vbCancel
79                m_numButsVisible = 2
80            Case vbAbortRetryIgnore
81                m_but1Result = vbAbort
82                m_but2Result = vbRetry
83                m_but3Result = vbIgnore
84                m_numButsVisible = 3
85            Case vbYesNoCancel
86                m_but1Result = vbYes
87                m_but2Result = vbNo
88                m_but3Result = vbCancel
89                m_numButsVisible = 3
90            Case vbYesNo
91                m_but1Result = vbYes
92                m_but2Result = vbNo
93                m_numButsVisible = 2
94            Case vbRetryCancel
95                m_but1Result = vbRetry
96                m_but2Result = vbIgnore
97                m_numButsVisible = 2
98            Case Else
99                Throw ButtonsError
100       End Select

101       Select Case CStr(DefaultButtonNumber) & CStr(m_ButtonsSimple)
              Case "2" & CStr(vbOKOnly)
102               Throw "Illegal value for Buttons - Cannot specify both vbOKOnly and vbDefaultButton2"
103           Case "3" & CStr(vbOKOnly)
104               Throw "Illegal value for Buttons - Cannot specify both vbOKOnly and vbDefaultButton3"
105           Case "3" & CStr(vbOKCancel)
106               Throw "Illegal value for Buttons - Cannot specify both vbOKCancel and vbDefaultButton3"
107           Case "3" & CStr(vbYesNo)
108               Throw "Illegal value for Buttons - Cannot specify both vbYesNo and vbDefaultButton3"
109           Case "3" & CStr(vbRetryCancel)
110               Throw "Illegal value for Buttons - Cannot specify both vbRetryCancel and vbDefaultButton3"
111       End Select

          'Set fonts
112       For i = 1 To 5
113           Set Ctrl = Choose(i, Me.but1, Me.but2, Me.but3, Me.CheckBox1, Me.TextBox1)
114           Ctrl.Font.Name = m_FontName
115           Ctrl.Font.Size = m_FontSize
116       Next i

          'Set captions
117       If Caption1 = vbNullString Then
118           Caption1 = Choose(m_but1Result, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No", "Help")
119       End If
120       If m_numButsVisible >= 2 Then
121           If Caption2 = vbNullString Then
122               Caption2 = Choose(m_but2Result, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No", "Help")
123           End If
124       End If
125       If m_numButsVisible >= 3 Then
126           If Caption3 = vbNullString Then
127               Caption3 = Choose(m_but3Result, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No", "Help")
128           End If
129       End If

          'Set values and visible status of controls
          Dim Accelerator As String
130       m_AcceleratorsUsed = vbNullString
131       Me.caption = Title
132       Me.but1.caption = ProcessCaption(Caption1, Accelerator)
133       Me.but1.Accelerator = Accelerator
134       Me.but2.caption = ProcessCaption(Caption2, Accelerator)
135       Me.but2.Accelerator = Accelerator
136       Me.but3.caption = ProcessCaption(Caption3, Accelerator)
137       Me.but3.Accelerator = Accelerator

138       For i = 1 To 3
139           Me.Controls("but" + CStr(i)).Visible = i <= m_numButsVisible
140       Next i

141       If CheckBoxCaption <> vbNullString Then
142           Me.CheckBox1.Visible = True
143           Me.CheckBox1.caption = ProcessCaption(CheckBoxCaption, Accelerator)
144           Me.CheckBox1.Accelerator = Accelerator
145           Me.CheckBox1.Value = CheckBoxValue
146       Else
147           Me.CheckBox1.Visible = False
148       End If

149       If m_but2Result = vbCancel Then Me.but2.Cancel = True
150       If m_but3Result = vbCancel Then Me.but3.Cancel = True
151       If m_ButtonsSimple = vbOKOnly Or m_ButtonsSimple = vbOKOnly + vbMsgBoxHelpButton Then
152           Me.but1.Cancel = True
153       End If

          Dim x As Double
          Dim y As Double
154       x = m_GapOnLeft        ' x and y are the "write coordinates"
155       y = m_GapAboveLabel

          'Position the Image
156       If Not (imgToShow Is Nothing) Then
157           imgToShow.Left = x: imgToShow.Top = y
158           x = x + imgToShow.Width + 5
159       End If

          'Position and Size TextBox1
160       With Me.TextBox1
161           Me.TextBox1.Value = Prompt + vbLf + "xxx" 'PGS 23 July 2020. Was finding Autosize not working correctly on very high dpi setups. So temporarily add an extra line to the prompt...
162           .Top = y
163           .Left = x
164           .Width = 10000
165           .AutoSize = False
166           .Width = TextWidth
167           .AutoSize = True
168           .AutoSize = False
              'PGS 22 Jan 2022. AutoSize still proving to be troublesome. Better the dialog is too big than too small, so boost a bit
169           .Width = .Width + 10
170           .Height = .Height + 10
171           Me.TextBox1.Value = Prompt

172           If .Height > m_MaxLabelHeight Then
173               .ScrollBars = fmScrollBarsNone        'scroll bars need to be kicked into life
174               .ScrollBars = fmScrollBarsVertical
175               .Height = m_MaxLabelHeight
176               .Width = .Width + 20        'make room for scroll bars
177               .SetFocus
178               .SelStart = 0
179               .SelLength = 0
180           End If

              'Try to make the roll-our-own hyperlinks discoverable.
181           If InStr(Prompt, "www.") > 0 Or InStr(Prompt, "http://") > 0 Or InStr(Prompt, "https://") > 0 Then
182               .ControlTipText = "Click web-addresses to go there."
183           End If

184           If Not (imgToShow Is Nothing) Then
185               If .Height < imgToShow.Height Then
                      'when there are many lines of text we align the tops of imgToShow and TextBox1, _
                       but when there's only one line of text we align the centres
186                   imgToShow.Top = .Top + .Height / 2 - imgToShow.Height / 2
187               End If
188           End If
189       End With

190       With Me.Controls("but" & CStr(DefaultButtonNumber))
191           .Default = True
192           .SetFocus
193       End With
          
          'Position and Size CheckBox1
194       If CheckBoxCaption <> vbNullString Then
195           x = m_GapOnLeft
196           y = Me.TextBox1.Top + Me.TextBox1.Height + m_GapAboveCheckBox
197           If Not imgToShow Is Nothing Then
198               y = SafeMax(y, imgToShow.Top + imgToShow.Height + m_GapAboveCheckBox)
199           End If

              Dim CheckboxTextWidth As Double
200           CheckboxTextWidth = 18 + sStringWidth(CheckBoxCaption, Me.CheckBox1.Font.Name, Me.CheckBox1.Font.Size)(1, 1)
201           Me.CheckBox1.Width = 10000
202           Me.CheckBox1.AutoSize = False
203           Me.CheckBox1.Width = TextWidth - m_CheckBoxNudge
204           Me.CheckBox1.AutoSize = True
205           Me.CheckBox1.AutoSize = False
206           Me.CheckBox1.Top = y
207           Me.CheckBox1.Left = Me.TextBox1.Left + m_CheckBoxNudge
208           x = m_GapOnLeft
209           y = Me.CheckBox1.Top + Me.CheckBox1.Height + m_GapAboveButtons
210       Else
211           x = m_GapOnLeft
212           y = Me.TextBox1.Top + Me.TextBox1.Height + m_GapAboveButtons
213           If Not imgToShow Is Nothing Then
214               y = SafeMax(y, imgToShow.Top + imgToShow.Height + m_GapAboveButtons)
215           End If
216       End If

217       Me.Frame1.Left = 0
218       Me.Frame1.Top = y

219       If SecondsToSelfDestruct > 0 Then
220           m_SecondsToSelfDestruct = SecondsToSelfDestruct
221           m_SelfDestructMode = True
222           m_SelfDestructButton = SelfDestructButton
223       Else
224           m_SelfDestructMode = False
225       End If

          'Position and Size Buttons.
          Dim MaxButHeight As Double
226       For i = 1 To m_numButsVisible
227           With Me.Controls("but" + CStr(i))
                  Dim origCaption As String

228               .WordWrap = True
229               origCaption = .caption
230               .caption = origCaption + "XX"        'PGS. Was finding that auto-size sometimes truncated last character or two. This "XX"! hack seems to cure the problem

231               If m_SelfDestructMode Then
232                   If Choose(i, m_but1Result, m_but2Result, m_but3Result) = m_SelfDestructButton Then
                          'get sizing correct with extra characters for timing
233                       .caption = origCaption + " (" & CStr(m_SecondsToSelfDestruct) + ")XX"
234                   End If
235               End If
236               .Width = 10000
237               .AutoSize = False
238               .Width = ButtonWidth
239               .AutoSize = True
240               .AutoSize = False
241               .caption = origCaption
242               If MaxButHeight < .Height Then MaxButHeight = .Height
243           End With
244       Next i
245       For i = 1 To m_numButsVisible
246           With Me.Controls("but" + CStr(i))
247               .Height = MaxButHeight
248               .Width = ButtonWidth
249               .Left = m_GapOnLeft + (i - 1) * (ButtonWidth + m_GapBetweenButtons)
250           End With
251       Next i

          'Me.Height-Me.InsideHeight is a measure of the height of the title bar at the top of the form, which changes in one version of Excel to the next...
252       Me.Height = Me.Frame1.Top + Me.but1.Top + Me.but1.Height + (Me.Height - Me.InsideHeight) + 11.35

          Dim maxLeft
253       maxLeft = Me.Controls("but" + CStr(m_numButsVisible)).Left + Me.Controls("but" + CStr(m_numButsVisible)).Width
254       maxLeft = SafeMax(maxLeft, Me.TextBox1.Left + Me.TextBox1.Width)
255       If CheckBoxCaption <> vbNullString Then maxLeft = SafeMax(maxLeft, Me.CheckBox1.Left + Me.CheckBox1.Width)

256       Me.Width = maxLeft + m_GapOnRight
257       Me.Frame1.Width = Me.Width + 10
258       Me.Frame1.Height = Me.Height - Me.Frame1.Top

          Dim NudgeRight
259       NudgeRight = Me.Width - (Me.Controls("but" + CStr(m_numButsVisible)).Left + Me.Controls("but" + CStr(m_numButsVisible)).Width) - m_GapOnRight

260       For i = 1 To m_numButsVisible
261           With Me.Controls("but" + CStr(i))
262               .Left = .Left + NudgeRight
263           End With
264       Next i

          'Set up resizer...
265       Me.but1.Tag = "L"
266       Me.but2.Tag = "L"
267       Me.but3.Tag = "L"
268       Me.Frame1.Tag = "TW"
269       Me.TextBox1.Tag = "WH"
270       CreateFormResizer mclsResizer
          Dim MinimumFormHeight As Double
          Dim ScreenHeight As Double
271       ScreenHeight = GetSystemMetrics(62) / fY()        'SM_CYMAXIMIZED, height of window when maximised
272       MinimumFormHeight = Me.Height
273       If MinimumFormHeight > ScreenHeight Then MinimumFormHeight = ScreenHeight
274       mclsResizer.Initialise Me, MinimumFormHeight, Me.Width, Me.Frame1.BackColor
275       If Me.Height > ScreenHeight Then
276           mclsResizer.ResizeControls Me, Me.Height, Me.Width, ScreenHeight, Me.Width
277       End If

278       Exit Sub
ErrHandler:
279       Throw "#frmMsgBoxPlus.Initialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

1         On Error GoTo ErrHandler
2         If (KeyCode = 67 And Shift = 2) Or (KeyCode = 45 And Shift = 2) Then        'Ctrl C or Ctrl Insert
3             If Me.TextBox1.SelLength > 0 Then
4                 CopyStringToClipboard Me.TextBox1.SelText
5             End If
6         End If

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#TextBox1_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         TextBoxMouseEvent Me.TextBox1, Button, Shift, x, y, True
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBox1_MouseDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         TextBoxMouseEvent Me.TextBox1, Button, Shift, x, y, False
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TextBox1_MouseUp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UserForm_Activate
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Purpose   : This code has to be in the Activate event, not the initialise routine,
'             otherwise the form never shows...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
1         If m_SelfDestructMode Then
2             FormSelfDestruct
3         End If
End Sub

Private Sub but1_Click()
1         On Error GoTo ErrHandler
2         m_ReturnValue = m_but1Result
3         m_ReturnCaption = Me.but1.caption
4         If Not m_SelfDestructMode Then
5             HideForm Me
6         Else
7             m_HideNow = True
8         End If
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#frmMsgBoxPlus.but1_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but2_Click()
1         On Error GoTo ErrHandler
2         m_ReturnValue = m_but2Result
3         m_ReturnCaption = Me.but2.caption
4         If Not m_SelfDestructMode Then
5             HideForm Me
6         Else
7             m_HideNow = True
8         End If
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#frmMsgBoxPlus.but2_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub but3_Click()
1         On Error GoTo ErrHandler
2         m_ReturnValue = m_but3Result
3         m_ReturnCaption = Me.but3.caption
4         If Not m_SelfDestructMode Then
5             HideForm Me
6         Else
7             m_HideNow = True
8         End If
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#frmMsgBoxPlus.but3_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#frmMsgBoxPlus.but1_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#frmMsgBoxPlus.but2_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#frmMsgBoxPlus.but3_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub CheckBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#frmMsgBoxPlus.CheckBox1_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#frmMsgBoxPlus.TextBox1_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         HighlightFormControl Me.but1
3         UnHighlightFormControl Me.but2
4         UnHighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.but1_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl Me.but1
3         HighlightFormControl Me.but2
4         UnHighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.but2_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub but3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl Me.but1
3         UnHighlightFormControl Me.but2
4         HighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.but3_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

1         On Error GoTo ErrHandler
2         UnHighlightFormControl Me.but1
3         UnHighlightFormControl Me.but2
4         UnHighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.Frame1_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler
2         UnHighlightFormControl Me.but1
3         UnHighlightFormControl Me.but2
4         UnHighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.TextBox1_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

1         On Error GoTo ErrHandler
2         UnHighlightFormControl Me.but1
3         UnHighlightFormControl Me.but2
4         UnHighlightFormControl Me.but3
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#frmMsgBoxPlus.UserForm_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UserForm_QueryClose
' Author    : Philip Swannell
' Date      : 21-Nov-2015
' Purpose   : User clicks red x in top corner, or hits escape key. Prevent dialog from
'             vanishing unless the user's intent is clear.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

1         On Error GoTo ErrHandler

2         If CloseMode = vbFormControlMenu Then
3             Select Case m_ButtonsSimple
                  Case vbAbortRetryIgnore, vbYesNo, vbOKOnly + vbMsgBoxHelpButton, vbAbortRetryIgnore + vbMsgBoxHelpButton, vbYesNo + vbMsgBoxHelpButton
4                     Cancel = 1
5                     Exit Sub
6                 Case vbOKOnly
7                     m_ReturnValue = vbOK
8                     m_ReturnCaption = Me.but1.caption
9                 Case vbOKCancel, vbRetryCancel, vbOKCancel + vbMsgBoxHelpButton, vbRetryCancel + vbMsgBoxHelpButton
10                    m_ReturnValue = vbCancel
11                    m_ReturnCaption = Me.but2.caption
12                Case vbYesNoCancel, vbYesNoCancel + vbMsgBoxHelpButton
13                    m_ReturnValue = vbCancel
14                    m_ReturnCaption = Me.but3.caption
15                Case Else
16                    Throw "Unrecognised value for Buttons"
17            End Select
18        End If

19        If m_SelfDestructMode Then
20            m_HideNow = True
21            Cancel = 1
22        End If

23        Exit Sub
ErrHandler:
24        Throw "#frmMsgBoxPlus.UserForm_QueryClose(line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ProcessCaption
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Purpose   : Strips & characters that indicate accelerator, and && that indicate &. Also
'             returns the accelerator key
' -----------------------------------------------------------------------------------------------------------------------
Private Function ProcessCaption(caption As String, ByRef Accelerator As String)
          Dim i As Long
          Dim ThisChar As String
1         On Error GoTo ErrHandler
2         Accelerator = vbNullString

3         ProcessCaption = ProcessAmpersands(caption, Accelerator)
4         If Accelerator = vbNullString Then
5             For i = 1 To Len(ProcessCaption)
6                 ThisChar = UCase$(Mid$(ProcessCaption, i, 1))
7                 If InStr(m_AcceleratorsUsed, ThisChar) = 0 Then
8                     If ThisChar <> " " Then        '" " character is a bad choice for accelerator, especially for check-boxes for which space bar on the keyboard is the accelerator if that check box has focus
9                         Accelerator = ThisChar
10                        Exit For
11                    End If
12                End If
13            Next i
14        End If

15        m_AcceleratorsUsed = m_AcceleratorsUsed + Accelerator
16        Exit Function
ErrHandler:
17        Throw "#frmMsgBoxPlus.ProcessCaption (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : KeyPressResponse
' Author    : Philip Swannell
' Date      : 15-Nov-2015
' Purpose   : The buttons and checkbox have accelerator keys. i.e. the button gets focus
'             if the user hits Alt + character. But this routine makes it the case that simply
'             hitting the accelerator key (without Alt) is equivalent to clicking that button\checkbox.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub KeyPressResponse(KeyAscii As MSForms.ReturnInteger)
          Dim c As control
          Dim i As Long
          Dim N As Long
          Dim NumberChosen As Long

1         On Error GoTo ErrHandler
2         For i = 1 To 3
3             Set c = Me.Controls("but" & CStr(i))
4             If Not c.Visible Then Exit For
5             If UCase$(c.Accelerator) = UCase$(Chr$(KeyAscii)) Then
6                 N = N + 1
7                 NumberChosen = i
8                 If N > 1 Then Exit Sub        'more than one button has the same accelerator
9             End If

10        Next i
11        If N = 1 Then
12            Select Case NumberChosen
                  Case 1
13                    m_ReturnValue = m_but1Result
14                    m_ReturnCaption = Me.but1.caption
15                Case 2
16                    m_ReturnValue = m_but2Result
17                    m_ReturnCaption = Me.but2.caption
18                Case 3
19                    m_ReturnValue = m_but3Result
20                    m_ReturnCaption = Me.but3.caption
21            End Select
22            If Not m_SelfDestructMode Then
23                HideForm Me
24            Else
25                m_HideNow = True
26            End If
27        End If
28        If N = 0 Then
29            If Me.CheckBox1.Visible Then
30                If UCase$(Me.CheckBox1.Accelerator) = UCase$(Chr$(KeyAscii)) Then
31                    Me.CheckBox1.SetFocus
32                    Me.CheckBox1.Value = Not (Me.CheckBox1.Value)
33                End If
34            End If
35        End If
36        Exit Sub
ErrHandler:
37        Throw "#frmMsgBoxPlus.KeyPressResponse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetSelfDestructCaption
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Purpose   : Updates the caption on the "self-destruct" button. Called from FormSelfDestruct
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SetSelfDestructCaption(SelfDestructButton As VbMsgBoxResult, SecondsRemaining As Long)

          Static HaveCalledBefore As Boolean
          Static OriginalCaption As String
          Static TheButtonToChange As control

1         On Error GoTo ErrHandler
2         If Not HaveCalledBefore Then
3             If m_but1Result = SelfDestructButton Then
4                 Set TheButtonToChange = Me.but1
5             ElseIf Me.but2.Visible And SelfDestructButton = m_but2Result Then
6                 Set TheButtonToChange = Me.but2
7             ElseIf Me.but3.Visible And SelfDestructButton = m_but3Result Then
8                 Set TheButtonToChange = Me.but3
9             Else
10                Throw "SelfDestructButton does not correspond to a button displayed"
11            End If
12            OriginalCaption = TheButtonToChange.caption
13        End If
14        HaveCalledBefore = True

15        TheButtonToChange.caption = OriginalCaption & " (" & SecondsRemaining & ")"
16        Exit Sub
ErrHandler:
17        Throw "#frmMsgBoxPlus.SetSelfDestructCaption (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormSelfDestruct
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Purpose   : In "self destruct mode" (m_SelfDestructMode = True) it's always this routine
'             that dismisses the form. Clicking on one of the buttons simply sets the m_HideNow
'             flag so that this method does its work immediately.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FormSelfDestruct()
          Dim PrevSecondsLeft
          Dim SecondsLeft As Long
          Dim TimeNow As Double

1         On Error GoTo ErrHandler

2         TimeNow = sElapsedTime()

3         PrevSecondsLeft = 0
4         Do While sElapsedTime() < TimeNow + m_SecondsToSelfDestruct
5             If m_HideNow Then Exit Do
6             SecondsLeft = CLng(m_SecondsToSelfDestruct - (sElapsedTime() - TimeNow))
7             If (SecondsLeft <> PrevSecondsLeft) And (SecondsLeft <> 0) Then
8                 PrevSecondsLeft = SecondsLeft
9                 SetSelfDestructCaption m_SelfDestructButton, SecondsLeft
10            End If
11            DoEvents
12        Loop
13        If Not m_HideNow Then
14            m_ReturnValue = m_SelfDestructButton
15        End If
16        HideForm Me

17        Exit Sub

ErrHandler:
18        Throw "#frmMsgBoxPlus.FormSelfDestruct (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

