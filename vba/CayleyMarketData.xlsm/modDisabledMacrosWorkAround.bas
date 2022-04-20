Attribute VB_Name = "modDisabledMacrosWorkAround"
Option Explicit

'PGS 14 April 2022

'Have started seeing problem where buttons on sheets of this workbook whose OnAction "points to" SolumAddin won't work, clicking on them
'displays a message box such as "Cannot run the macro 'GroupingButton'. The macro may not be available in this workbook or all macros may be disabled"

'Have tried various work arounds:
'1) Reconstructing the workbook by copying all sheets, modules, references etc to a new workbook. The problem goes away until the new workbook is saved
'and re-opened at which point the problem comes back.
'2) Ensuring that SolumAddin is saved at a Trusted Location.

'I discovered that flipping the IsAddin property of SolumAddin.xlam to False makes the problem go away. But that's obvs. not a solution.

'So not very satisfactory solution is to wrap methods in SomAddin with methods in this module and assign the wrappers to the buttons.

Sub ReassignButtonOnActions()
          Dim b As Button
          Dim ws As Worksheet

          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             Set SPH = CreateSheetProtectionHandler(ws)
4             For Each b In ws.Buttons
5                 Select Case SafeOnAction(b)
                      Case "GroupingButton", "SAISortButtonOnAction"
6                         b.OnAction = b.OnAction & "Wrap"
7                     Case "SolumAddin.xlam!'SAISortButtonOnAction 3 '"
8                         b.OnAction = "SAISortButtonOnActionWrap3"
9                 End Select
10            Next
11        Next

12        Exit Sub

ErrHandler:
13        Throw "#ReassignButtonOnActions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function SafeOnAction(b As Button)

1         On Error GoTo ErrHandler
2         SafeOnAction = b.OnAction
3         Exit Function
ErrHandler:
4         SafeOnAction = "Can't get the OnAction FFS"
End Function

Sub GroupingButtonWrap()
1         GroupingButton
End Sub

Sub SAISortButtonOnActionWrap()
1         SAISortButtonOnAction 1
End Sub

Sub SAISortButtonOnActionWrap3()
1         SAISortButtonOnAction 3
End Sub

Sub AuditMenuWrap()
1         AuditMenu
End Sub

