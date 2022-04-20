Attribute VB_Name = "modInsertFunction"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modInsertFuntion
' Author    : Philip Swannell
' Date      : 25-Oct-2018
' Purpose   : Common code between "Function Vrowser" and ribbon menus on the
'             'Solum' > 'Worksheet Function' group
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsExcelReadyForFormula
' Author    : Philip Swannell
' Date      : 18-Nov-2013
' Purpose   : Encapsulate decision of whether or not Excel is in a good state for the
'             formula to be entered.
'             Argument EnteringArrayFormula can be True, False or Missing meaning we don't
'             yet know what the user will choose to do.
' -----------------------------------------------------------------------------------------------------------------------
Function IsExcelReadyForFormula(SilentMode As Boolean, Optional EnteringArrayFormula As Variant) As Boolean

          Dim Prompt As String

1         On Error GoTo ErrHandler

2         If ActiveSheet Is Nothing Then
3             Prompt = "There is no active worksheet."
4         ElseIf TypeName(ActiveSheet) <> "Worksheet" Then
5             Prompt = "Active sheet must be a worksheet"
6         ElseIf ActiveWindow Is Nothing Then
7             Prompt = "There is no active window."
8         ElseIf TypeName(Selection) <> "Range" Then
9             Prompt = "No cells are selected."
10        ElseIf ActiveSheet.ProtectContents Then
11            If SheetIsProtectedWithPassword(ActiveSheet) Then
12                Prompt = "You cannot use this command on a protected sheet. To use this" & _
                      " command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button)." & _
                      " You will be prompted for a password."
13            Else
14                Prompt = "You cannot use this command on a protected sheet. To use this" & _
                      " command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button)."
15            End If
16        ElseIf Selection.Areas.Count > 1 Then
17            Prompt = "That command cannot be used on multiple selections."
18        ElseIf VarType(EnteringArrayFormula) = vbBoolean Then
19            If EnteringArrayFormula Then
20                If ActiveCell.HasArray Then
21                    If ActiveCell.address <> ActiveCell.CurrentArray.Cells(1, 1).address Then
22                        Prompt = "To replace the array formula at " + AddressND(ActiveCell.CurrentArray) + vbLf + _
                              "please first select its top-left cell " + AddressND(ActiveCell.CurrentArray.Cells(1, 1)) + "."
23                    End If
24                End If
25            Else
26                If ActiveCell.HasArray Then
27                    If ActiveCell.CurrentArray.Cells.CountLarge > 1 Then
28                        Prompt = "You cannot change part of the array at " & _
                              AddressND(ActiveCell.CurrentArray) & "." + vbLf + vbLf + _
                              "Hint:" + vbLf + "Ctrl Shift Enter will replace an existing array formula with a new one."
29                    End If
30                End If
31            End If
32        End If

33        If Prompt = vbNullString Then
34            IsExcelReadyForFormula = True
35        ElseIf Not SilentMode Then
36            MsgBoxPlus Prompt, vbExclamation, "Insert " & gAddinName & " function"
37        End If

38        Exit Function
ErrHandler:
39        Throw "#IsExcelReadyForFormula (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ConfirmInsertFunction
' Author    : Philip Swannell
' Date      : 18-Nov-2013
' Purpose   : Ask the user if their sure they want to overwrite cells already on the sheet.
' -----------------------------------------------------------------------------------------------------------------------
Function ConfirmInsertFunction(EnteringArrayFormula As Boolean, ChosenFunction As String) As Boolean
          Dim Prompt As String

1         On Error GoTo ErrHandler

2         If EnteringArrayFormula Then
3             If Not IsEmpty(ActiveCell.Value) Then
4                 If ActiveCell.HasArray Then
5                     Prompt = "Replace array formula:" + vbLf + vbLf + _
                          "{" + ActiveCell.Formula + "}" + vbLf + vbLf + _
                          "in cell" + IIf(ActiveCell.CurrentArray.Cells.CountLarge > 1, "s", vbNullString) + " " + AddressND(ActiveCell.CurrentArray) + " with an array formula:" + vbLf + vbLf + _
                          "{=" + ChosenFunction + "(...)}?"
6                 ElseIf ActiveCell.HasFormula Then
7                     Prompt = "Replace formula:" + vbLf + vbLf + _
                          ActiveCell.Formula + vbLf + vbLf + _
                          "in cell " + AddressND(ActiveCell) + " with an array formula:" + vbLf + vbLf + _
                          "{=" + ChosenFunction + "(...)}?"
8                 Else
9                     Prompt = "Overwrite cell " + AddressND(ActiveCell) + " containing:" + vbLf + vbLf + _
                          ActiveCell.text + vbLf + vbLf + _
                          "with an array formula:" + vbLf + vbLf + _
                          "{=" + ChosenFunction + "(...)}?"
10                End If
11            End If
12        Else        'Entering non-array formula
13            If Not IsEmpty(ActiveCell.Value) Then
14                If ActiveCell.HasArray Then
15                    Prompt = "Replace array formula:" + vbLf + vbLf + _
                          "{" + ActiveCell.Formula + "}" + vbLf + vbLf + _
                          "in cell " + AddressND(ActiveCell.CurrentArray) + " with a formula:" + vbLf + vbLf + _
                          "=" + ChosenFunction + "(...)?"

16                ElseIf ActiveCell.HasFormula Then
17                    Prompt = "Replace formula:" + vbLf + vbLf + _
                          ActiveCell.Formula + vbLf + vbLf + _
                          "in cell " + AddressND(ActiveCell) + " with a formula:" + vbLf + vbLf + _
                          "=" + ChosenFunction + "(...)?"
18                Else
19                    Prompt = "Overwrite cell " + AddressND(ActiveCell) + " containing:" + vbLf + vbLf + _
                          ActiveCell.text + vbLf + vbLf + _
                          "with a formula:" + vbLf + vbLf + _
                          "=" + ChosenFunction + "(...)?"
20                End If
21            End If
22        End If

23        If Prompt = vbNullString Then
24            ConfirmInsertFunction = True
25        Else
26            ConfirmInsertFunction = MsgBoxPlus(Prompt, vbYesNo + vbDefaultButton2 + vbQuestion, IIf(EnteringArrayFormula, "Insert Array Formula", "Insert Formula")) = vbYes
27        End If

28        Exit Function
ErrHandler:
29        Throw "#ConfirmInsertFunction (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InsertFunctionAtActiveCell
' Author    : Philip Swannell
' Date      : 24-Oct-2018
' Purpose   : Used from both the ribbon and the "function browser"
' -----------------------------------------------------------------------------------------------------------------------
Sub InsertFunctionAtActiveCell(ChosenFunction As String, Optional UseCtrlShiftEnter As Variant)

          Const MSGBOXTITLE = "Insert " & gAddinName & " function"
          Dim HaveChangedSheet As Boolean
          Dim MatchRes As Variant
          Dim NewFormula As String
          Dim NewSelection As Range
          Dim OrigActiveCell As Range
          Dim OrigSelection As Range
          Dim ResizeRes As rsafReturnValue
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         If VarType(UseCtrlShiftEnter) <> vbBoolean Then
3             MatchRes = sMatch(ChosenFunction, shHelp.Range("TheData").Columns(1).Value, True)
4             If IsNumber(MatchRes) Then
5                 UseCtrlShiftEnter = shHelp.Range("TheData").Cells(MatchRes, 0).Value
6             Else
7                 UseCtrlShiftEnter = True
8             End If
9         End If

10        If Not (IsExcelReadyForFormula(False, UseCtrlShiftEnter)) Then Exit Sub
11        If Not ConfirmInsertFunction(CBool(UseCtrlShiftEnter), ChosenFunction) Then Exit Sub

12        Set OrigSelection = Selection
13        Set OrigActiveCell = ActiveCell
14        Set ws = OrigSelection.Parent
15        BackUpRange OrigSelection.CurrentRegion, shUndo2, OrigSelection

          'If the function selected is already in use in the active cell
16        If Left$(ActiveCell.Formula, Len(ChosenFunction) + 2) = "=" & ChosenFunction & "(" Then
17            If ActiveCell.HasArray And UseCtrlShiftEnter Or Not ActiveCell.HasArray And Not UseCtrlShiftEnter Then
18                GoTo ShowWizard
19            ElseIf Not ActiveCell.HasArray And UseCtrlShiftEnter Then
20                If Not SetFormulaArray(ActiveCell, ActiveCell.Formula) Then
21                    RestoreRangeFromUndoBuffer2
22                    MsgBoxPlus "Formula is too long.", vbExclamation, MSGBOXTITLE
23                    Exit Sub
24                Else
25                    GoTo ShowWizard
26                End If
27                Exit Sub
28            Else        ' i.e. ActiveCell.HasArray And Not UseCtrlShiftEnter
29                If ActiveCell.CurrentArray.Cells.CountLarge > 1 Then
30                    MsgBoxPlus "You cannot change part of an array.", vbExclamation, MSGBOXTITLE
31                    Exit Sub
32                Else
33                    ActiveCell.FormulaR1C1 = ActiveCell.FormulaR1C1: HaveChangedSheet = True
34                End If
35            End If
36        Else        'Function selected is not already in use in the active cell.
37            NewFormula = "=" + ChosenFunction + "()"
38            If ActiveCell.HasArray And UseCtrlShiftEnter Then
39                Set NewSelection = ActiveCell.CurrentArray
40                BackUpRange NewSelection, shUndo
41                NewSelection.Select
42                NewSelection.FormulaArray = NewFormula: HaveChangedSheet = True
43                GoTo ShowWizard
44            ElseIf Not ActiveCell.HasArray And UseCtrlShiftEnter Then
45                ActiveCell.FormulaArray = NewFormula: HaveChangedSheet = True
46                GoTo ShowWizard
47            ElseIf ActiveCell.HasArray And Not UseCtrlShiftEnter Then
48                If ActiveCell.CurrentArray.Cells.CountLarge > 1 Then
49                    MsgBoxPlus "You cannot change part of an array.", vbExclamation, MSGBOXTITLE
50                    Exit Sub
51                Else
52                    If ExcelSupportsSpill() Then
53                        ActiveCell.Formula2 = NewFormula: HaveChangedSheet = True
54                    Else
55                        ActiveCell.Formula = NewFormula: HaveChangedSheet = True
56                    End If
57                    GoTo ShowWizard
58                End If
59            Else        'i.e. Not ActiveCell.HasArray And Not UseCtrlShiftEnter
60                If ExcelSupportsSpill() Then
61                    ActiveCell.Formula2 = NewFormula: HaveChangedSheet = True
62                Else
63                    ActiveCell.Formula = NewFormula: HaveChangedSheet = True
64                End If
65                GoTo ShowWizard
66            End If
67        End If

ShowWizard:
68        If Not Application.Dialogs(xlDialogFunctionWizard).Show Then
              'User has hit Cancel in Excel Function Wizard
69            If HaveChangedSheet Then
70                RestoreRangeFromUndoBuffer2
71            End If
72            Exit Sub
73        ElseIf UseCtrlShiftEnter Then
74            OrigSelection.Select
75            ResizeRes = FitArrayFormula()
76            If ResizeRes = rsafAborted Or ResizeRes = rsafFailedButCorrectCleanup Then        'What about the other potential returns from  ResizeArrayFormula???
77                RestoreRangeFromUndoBuffer2
78                Exit Sub
79            End If
80            If IsUndoAvailable(shUndo) And IsUndoAvailable(shUndo2) Then
81                Application.OnUndo "Undo Insert " & gAddinName & " Function", ThisWorkbook.Name & "!RestoreRangeTwice"
82            End If
83            Exit Sub
84        End If

85        If HaveChangedSheet Then
86            If IsUndoAvailable(shUndo2) Then
87                Application.OnUndo "Undo Insert " & gAddinName & " Function", ThisWorkbook.Name & "!RestoreRangeFromUndoBuffer2"
88            End If
89        End If

90        Exit Sub
ErrHandler:
91        SomethingWentWrong "#InsertFunctionAtActiveCell (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, MSGBOXTITLE
End Sub
