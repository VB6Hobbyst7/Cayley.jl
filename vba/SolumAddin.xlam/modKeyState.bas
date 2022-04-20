Attribute VB_Name = "modKeyState"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modKeyState
' Author    : Philip Swannell
' Date      : 06-Oct-2015
' Purpose   : Three functions to test if (either) Shift, Ctrl or Alt keys are pressed.
'             This is simplified version of Chip Pearson's code at http://www.cpearson.com/excel/keytest.aspx
'             Pearson's version can distinguish between the two Shift keys, but I can't imagine wanting to do that...
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text
Private Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer

Private Const KEY_MASK As Integer = &HFF80        ' decimal -128

Public Function IsShiftKeyDown() As Boolean
1         IsShiftKeyDown = CBool(GetKeyState(vbKeyShift) And KEY_MASK)
End Function

Public Function IsControlKeyDown() As Boolean
1         IsControlKeyDown = CBool(GetKeyState(vbKeyControl) And KEY_MASK)
End Function

Public Function IsAltKeyDown() As Boolean
1         IsAltKeyDown = CBool(GetKeyState(vbKeyMenu) And KEY_MASK)
End Function
