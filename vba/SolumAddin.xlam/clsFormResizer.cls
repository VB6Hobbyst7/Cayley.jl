VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsFormResizer
' Author    : Philip Swannell
' Date      : 13-Nov-2013
' Remarks   : I have found two classes on the web for making VBA User forms resizable.
'             The first is due to Stephen Bullen (www.oaltd.co.uk) and Rob Bovey (www.appspro.com)
'             and the second to Andy Pope (http://www.andypope.info/vba/resizeform.htm)
'             This class is my melding together of the two, taking the features of each that I
'             liked. One advantage of this code is that it makes no Windows API calls and
'             therefore should work on all versions of Windows without change.
'             The code places a "Drag Handle" at the bottom right of the form and the user can drag this
'             to resize the form and its controls. Unlike the Bullen/Bovey version the user cannot
'             drag the borders of the form - only the corner.
'             The code makes use of Tag properties of the parent form's controls to decide
'             which controls should change position or shape  - allowed letters in the Tag are TLWH
'             for Top, Left, Width and Height. If T is in the Tag then the Top of the control will move
'             as the form's height is changed. Likewise for L,W and H. Further syntax allowed is say T0.5L0.5
'             in which case the control will move half the amount of the change in the form's size.
'Usage     1) The Tag properties of the controls of the form must be set, either at design time or run time
'             in the initialize method of the form.
'          2) The code of the form should contain the following:
'             Private m_clsResizer As clsFormResizer
'             Private Sub UserForm_Initialize()
'                 'Put any code that sets the form's start up size here
'                 Set m_clsResizer = New clsFormResizer
'                 m_clsResizer.Add Me
'             End Sub
'             Private Sub UserForm_Terminate()
'                 Set m_clsResizer = Nothing
'             End Sub
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

Private WithEvents m_objResizer As MSForms.Frame
Attribute m_objResizer.VB_VarHelpID = -1
Private m_MouseX As Double
Private m_MouseY As Double
Private m_blnResizing As Boolean
Private WithEvents m_frmParent As MSForms.UserForm
Attribute m_frmParent.VB_VarHelpID = -1
Private m_objParent As Object
Private m_frmHeight As Double
Private m_frmWidth As Double
Private m_frmMinHeight As Double
Private m_frmMinWidth As Double
Const m_MsgBoxTitle = "clsFormResizer"
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Philip Swannell, adapted from Andy Pope
' Date      : 13-Nov-2013
' Purpose   : Tell the class what form it's handling add a resizing control to bottom
'             righthand corner.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Initialise(Parent As Object, Optional MinimumHeight As Double = 99, Optional MinimumWidth As Double = 28.5, Optional GrabHandleBackColor)

          Dim labTemp As MSForms.Label

1         On Error GoTo ErrHandler

2         Set m_frmParent = Parent
3         Set m_objParent = Parent
4         m_frmHeight = Parent.Height
5         m_frmWidth = Parent.Width
6         m_frmMinHeight = IIf(MinimumHeight < 99, 99, MinimumHeight)
7         m_frmMinWidth = IIf(MinimumWidth < 28.5, 28.5, MinimumWidth)

8         Set m_objResizer = m_objParent.Controls.Add("Forms.Frame.1", "FrameResizeGrab", True)
9         Set labTemp = m_objResizer.Add("Forms.label.1", "ResizeGrab", True)
10        With labTemp
11            With .Font
12                .Name = "Marlett"
13                .CharSet = 2
14                .Size = 14
15                .Bold = True
16            End With
17            .BackStyle = fmBackStyleTransparent
18            .AutoSize = True
19            .BorderStyle = fmBorderStyleNone
20            .caption = "o"
21            .MousePointer = fmMousePointerSizeNWSE
22            .ForeColor = RGB(100, 100, 100)
23            .ZOrder
24            .Top = 1
25            .Left = 1
26            .Enabled = False
27        End With

28        With m_objResizer
29            If Not IsMissing(GrabHandleBackColor) Then
30                .BackColor = GrabHandleBackColor
31            End If
32            .MousePointer = fmMousePointerSizeNWSE
33            .BorderStyle = fmBorderStyleNone
34            .SpecialEffect = fmSpecialEffectFlat
35            .ZOrder
36            .caption = vbNullString
37            .Width = labTemp.Width + 1
38            .Height = labTemp.Height + 1
39            .TabStop = False        '< PGS bug fix to Andy Pope's code
40            .Top = m_objParent.InsideHeight - .Height
41            .Left = m_objParent.InsideWidth - .Width
42        End With

43        Exit Sub
ErrHandler:
44        Throw "#Initialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : m_objResizer_MouseDown
' Author    : Philip Swannell, adapted from Andy Pope
' Date      : 13-Nov-2013
' Purpose   : Record where the mouse is when the user clicks on the GrabHandle
' -----------------------------------------------------------------------------------------------------------------------
Private Sub m_objResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         If Button = 1 Then
3             m_MouseX = x
4             m_MouseY = y
5             m_blnResizing = True
6         End If

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#m_objResizer_MouseDown (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : m_objResizer_MouseMove
' Author    : Philip Swannell, adapted from Andy Pope
' Date      : 13-Nov-2013
' Purpose   : resize the form and its controls as the user drags the drag handle.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub m_objResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
          Dim newHeight As Double
          Dim newWidth As Double
          Dim OldHeight As Double
          Dim oldWidth As Double

1         On Error GoTo ErrHandler

2         If Button = 1 Then
3             OldHeight = m_objParent.Height
4             oldWidth = m_objParent.Width
5             newWidth = oldWidth + x - m_MouseX
6             newHeight = OldHeight + y - m_MouseY
7             ResizeControls m_objParent, OldHeight, oldWidth, newHeight, newWidth
8         End If

9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#m_objResizer_MouseMove (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : m_objResizer_MouseUp
' Author    : Philip Swannell, adapted from Andy Pope
' Date      : 13-Nov-2013
' Purpose   : stop resizing the form
' -----------------------------------------------------------------------------------------------------------------------
Private Sub m_objResizer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1         On Error GoTo ErrHandler

2         If Button = 1 Then
3             m_blnResizing = False
4         End If

5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#m_objResizer_MouseUp (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : m_frmParent_Layout
' Author    : Philip Swannell
' Date      : 13-Nov-2013
' Purpose   : Handle the possibility that other code may change the size of the form and we need to
'             reposition all the controls.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub m_frmParent_Layout()
          Dim newHeight As Double
          Dim newWidth As Double
          Dim OldHeight As Double
          Dim oldWidth As Double

1         On Error GoTo ErrHandler

2         If Not m_blnResizing Then
3             OldHeight = m_frmHeight
4             oldWidth = m_frmWidth
5             newWidth = m_objParent.Width
6             newHeight = m_objParent.Height
7             ResizeControls m_objParent, OldHeight, oldWidth, newHeight, newWidth
8         End If

9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#m_frmParent_Layout (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeControls
' Author    : Philip Swannell, Adapted from Stephen Bullen / Rob Bovey
' Date      : 13-Nov-2013
' Purpose   : Change the position of the controls according to their Tags and the size
'             of the Form itself.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub ResizeControls(theForm As Object, OldHeight As Double, oldWidth As Double, newHeight As Double, newWidth As Double)

          Dim bSomeHeightChange As Boolean
          Dim bSomeWidthChange As Boolean
          Dim dHeightAdj As Double
          Dim dWidthAdj As Double
          Dim H As Double
          Dim L As Double
          Dim oCtl As MSForms.control
          Dim sTag As String
          Dim t As Double
          Dim w As Double

          'Calculate the change in height and width
1         On Error GoTo ErrHandler

          'Don't let the form get too small.
2         If newHeight < m_frmMinHeight Then newHeight = m_frmMinHeight
3         If newWidth < m_frmMinWidth Then newWidth = m_frmMinWidth

4         dHeightAdj = newHeight - OldHeight
5         dWidthAdj = newWidth - oldWidth

          'Check if we can perform the adjustment
          '(i.e. widths and heights can't be negative)
6         For Each oCtl In theForm.Controls

              'Read the control's Tag property, which contains the resizing info
7             sTag = UCase$(oCtl.Tag)

              'If we're changing the Top, check that it won't move off the top
              'of the form
8             If InStr(1, sTag, "T", vbBinaryCompare) Then
9                 If oCtl.Top + dHeightAdj * ResizeFactor(sTag, "T") <= 0 Then
10                    newHeight = OldHeight
11                End If
12                bSomeHeightChange = True
13            End If

              'If we're changing the Left, check that it won't move off the
              'left of the form
14            If InStr(1, sTag, "L", vbBinaryCompare) Then
15                If oCtl.Left + dWidthAdj * ResizeFactor(sTag, "L") <= 0 Then
16                    newWidth = oldWidth
17                End If
18                bSomeWidthChange = True
19            End If

              'If we're changing the Height, check that it won't go negative
20            If InStr(1, sTag, "H", vbBinaryCompare) Then
21                If oCtl.Height + dHeightAdj * ResizeFactor(sTag, "H") <= 0 Then
22                    newHeight = OldHeight
23                End If
24                bSomeHeightChange = True
25            End If

              'If we're changing the Width, check that it won't go negative
26            If InStr(1, sTag, "W", vbBinaryCompare) Then
27                If oCtl.Width + dWidthAdj * ResizeFactor(sTag, "W") <= 0 Then
28                    newWidth = oldWidth
29                End If
30                bSomeWidthChange = True
31            End If
32        Next        'Next OCtl

          'If none of the controls move or size,
          'don't allow the form to resize in that direction
33        If Not bSomeHeightChange Then newHeight = OldHeight
34        If Not bSomeWidthChange Then newWidth = oldWidth

          'Recalculate the height and width changes,
          'in case the previous checks reset them
35        dHeightAdj = newHeight - OldHeight
36        dWidthAdj = newWidth - oldWidth

37        If dHeightAdj <> 0 Or dWidthAdj <> 0 Then
              'Loop through all the controls on the form,
              'adjusting their position and size
38            For Each oCtl In theForm.Controls
39                With oCtl
40                    sTag = UCase$(.Tag)

41                    If dHeightAdj <> 0 Then
                          'Changing the Top
42                        If InStr(1, sTag, "T", vbBinaryCompare) Then
43                            t = .Top + dHeightAdj * ResizeFactor(sTag, "T")
44                            If .Top <> t Then .Top = t
45                        End If
                          'Changing the Height
46                        If InStr(1, sTag, "H", vbBinaryCompare) Then
47                            H = .Height + dHeightAdj * ResizeFactor(sTag, "H")
48                            If .Height <> H Then .Height = H
49                        End If
50                    End If

51                    If dWidthAdj <> 0 Then
                          'Changing the Left
52                        If InStr(1, sTag, "L", vbBinaryCompare) Then
53                            L = .Left + dWidthAdj * ResizeFactor(sTag, "L")
54                            If .Left <> L Then .Left = L
55                        End If
                          'Changing the Width
56                        If InStr(1, sTag, "W", vbBinaryCompare) Then
57                            w = .Width + dWidthAdj * ResizeFactor(sTag, "W")
58                            If .Width <> w Then .Width = w
59                        End If
60                    End If
61                End With
62            Next        'Next OCtl

63            If theForm.Height <> newHeight Then
64                theForm.Height = newHeight
65            End If
66            If theForm.Width <> newWidth Then
67                theForm.Width = newWidth
68            End If

69            With m_objResizer
70                .Left = theForm.InsideWidth - .Width
71                .Top = theForm.InsideHeight - .Height
72            End With

73            If newHeight <> OldHeight Or newWidth <> oldWidth Then
74                theForm.Repaint
75            End If
76            m_frmHeight = newHeight
77            m_frmWidth = newWidth

78        End If

79        Exit Sub
ErrHandler:
80        SomethingWentWrong "#ResizeControls (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, m_MsgBoxTitle
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeFactor
' Author    : Philip Swannell, adapted from Stephen Bullen/Rob Bovey
' Date      : 13-Nov-2013
' Purpose   : Interpret the Tags
' -----------------------------------------------------------------------------------------------------------------------
Private Function ResizeFactor(sTag As String, sChange As String)

          Dim dfactor As Double
          Dim iPos As Long

          'Locate the property letter in the tag string
1         On Error GoTo ErrHandler

2         iPos = InStr(1, sTag, sChange, vbBinaryCompare)

          'If we found it...
3         If iPos > 0 Then

              '... read the number following it
4             dfactor = Val(Mid$(sTag, iPos + 1))

              'If there was no number, use a factor of 100%, but cope with case where number is explicitly zero
5             If dfactor = 0 Then
6                 If Mid$(sTag, iPos + 1) <> "0" Then
7                     dfactor = 1
8                 End If
9             End If

10        End If

          'Return the factor
11        ResizeFactor = dfactor

12        Exit Function
ErrHandler:
13        Throw "#ResizeFactor (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Philip Swannell, Adapted from Andy Pope
' Date      : 13-Nov-2013
' Purpose   : Remove the Drag Handle if the class is terminated.
'             PGS 28/3/16 Have found that the attempt to remove the grab handle can cause
'             an Excel crash if triggered by the UserForm_QueryClose event of the parent
'             form. I can see no good reason to have this method anyway.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
1         On Error GoTo ErrHandler

          '  m_objParent.Controls.Remove "ResizeGrab"

2         Exit Sub
ErrHandler:
3         Throw "#Class_Terminate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

