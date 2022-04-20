VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsOptionButton
' Author    : Philip Swannell
' Date      : 15-Dec-2015
' Purpose   : Event handling code for the Option Buttons in frmOptionButtons that are
'             created at run time
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Public WithEvents butEvents As MSForms.OptionButton
Attribute butEvents.VB_VarHelpID = -1
Public Tag As Long

Private Sub butEvents_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
1         On Error GoTo ErrHandler
2         If butEvents.Parent.Parent.m_NumGroups = 1 Then
3             Cancel.Value = True
4             butEvents.Parent.Parent.Clicked_OK
5         End If
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#butEvents_DblClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butEvents_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
1         On Error GoTo ErrHandler
2         butEvents.Parent.Parent.KeyPressResponse KeyAscii
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butEvents_KeyPress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub butEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
1         On Error GoTo ErrHandler
2         butEvents.Parent.Parent.KeyDownResponse KeyCode, Shift
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#butEvents_KeyDown (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

