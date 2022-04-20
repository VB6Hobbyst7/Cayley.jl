VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScreenUpdateHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsScreenUpdateHandler
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Class to make handling of Application.ScreenUpdating easy and avoid screen
'             flicker. In general we want ScreenUpdating to be False while macros run, and
'             we also want the Application.Cursor to be xlWait, as an indication that code
'             is running.
'             Call method CreateScreenUpdateHandler() at the start of Top-Level routines.
'             A common use case is at the start of a method that displays a dialog to call
'             CreateScreenUpdateHandler(True) to set screen updating to True and cursor to
'             xlDefault for the duration of that method.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_OrigState As Boolean
Private m_origCursor As Variant

Private Sub Class_Initialize()
1         m_OrigState = Application.ScreenUpdating
2         m_origCursor = Application.Cursor
End Sub

Public Sub Init(SetStateTo As Boolean)
1         If Application.ScreenUpdating <> SetStateTo Then
2             Application.ScreenUpdating = SetStateTo
3         End If
4         If SetStateTo = True Then
5             If Application.Cursor <> xlDefault Then Application.Cursor = xlDefault
6         Else
7             If Application.Cursor <> xlWait Then Application.Cursor = xlWait
8         End If
End Sub

Private Sub Class_Terminate()
1         If Application.ScreenUpdating <> m_OrigState Then Application.ScreenUpdating = m_OrigState
2         If Application.Cursor <> m_origCursor Then Application.Cursor = m_origCursor
End Sub
