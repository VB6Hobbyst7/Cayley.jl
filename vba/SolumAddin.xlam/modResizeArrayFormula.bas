Attribute VB_Name = "modResizeArrayFormula"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modResizeArrayFormula
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Implements resizing of array formulae - either resizing the formula at the
'             active cell to the current selection or (in Auto mode) resizing to the "correct" size.
'Complexities:
'          1) Range.ArrayFormula = Formula fails if Formula has more than 255 characters.
'             We get around this by using SendKeys. SendKeys is hard to work with since a) it doesn't
'             work during a debugging session and b) the keys are only processed by Excel once macro
'             execution halts (I tried using DoEvents without success). Therefore we have to use Application.OnTime
'             to resume macro execution after a delay of one second.
'          2) The obvious approach to finding out what the "right size" of an array formula
'             is to use Application.Evaluate. Unfortunately this is quite unreliable. Sometimes
'             the call to Application.Evaluate throws an error, sometimes it works but evaluates
'             to an array of the wrong size - (e.g. for formulas using various native Excel functions such as INDEX)
'             Hence I use an alternative approach to determine the "Correct" size of the array. This is done by
'             pasting an amended version of the existing formula into a "spare" pair of cells of the spreadsheet. The
'             amendment is simply to wrap the existing formula with a call to sSizeOf which returns the
'             number of rows and columns in an array. Additional complexity: That formula must be entered as
'             an array formula into more than one cell - this solves the INDEX problem referred to above.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module
Private m_DoneBackup As Boolean
Private m_CurrentFormula As String
Private m_CurrentSizingFormula As String
Private m_Fit As Boolean
Private Const m_Title = "Resize Array Formula"
Private m_TempCellsAddress As String
Private m_ActiveCellAddress As String
Private m_CurrentRange As Range
Private m_RangeToSelectAfterError As Range
Private m_NewRange As Range
Private m_VisibleRangeAddress As String
Private m_ExcelState As clsExcelStateHandler
Private m_originalSelection As Range
Private m_origScrollArea As String
Private Const MAX_CELLS_FOR_RESIZE = 10000000#

Enum rsafResumePoint
    rsafResumeNone = 0
    rsafResumeAfterFirstSendKeys = 1
    rsafResumeAfterSecondSendKeys = 2
End Enum

Enum rsafReturnValue
    rsafFailed = 0
    rsafAborted = 1
    rsafFailedButCorrectCleanup = 2
    rsafNotFinishedAwaitingSendKeys = 3
    rsafWorkedCorrectly = 4
End Enum
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FitArrayFormula
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Resizes an array formula at the active cell to be the "right size" i.e. the
'             same number of rows and columns as the array being returned by the array
'             formula. Attached to key Ctrl+Shift+R and available from the Ribbon.
' -----------------------------------------------------------------------------------------------------------------------
Function FitArrayFormula()
1         On Error GoTo ErrHandler
2         If ExcelSupportsSpill() Then 'PGS 1 Aug 2019 - preparing for Dynamic Array Formulas, which really make the functionality of this module obsolete.
3             FitArrayFormula2
4             FitArrayFormula = rsafWorkedCorrectly
5         Else
6             FitArrayFormula = ResizeArrayFormula(True)
7         End If

8         Exit Function
ErrHandler:
9         SomethingWentWrong "#FitArrayFormula (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeArrayFormula
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Resizes an array formula at the active cell to be the size of the current
'             selection, or the "right size" if Fit is True. Available from the
'             Ribbon and from Ctrl+Shift+A.
'             There are two points in the code at which we sometimes have to resort to use
'             of SendKeys. In such cases we exit the function so as to to allow Excel time
'             to process the keys and code execution is restarted via Application.OnTime.
'             Hence the argument ResumePoint and the use of GoTo.
' -----------------------------------------------------------------------------------------------------------------------
Function ResizeArrayFormula(Optional Fit As Boolean, _
        Optional ResumePoint As rsafResumePoint = rsafResumeNone) As rsafReturnValue
          Dim NumCols As Long
          Dim NumRows As Long
          Dim Prompt As String
          Dim RangeToTest As Range
          Dim StatusBarText As String
          Dim TempCells As Range
          Dim TempString As String

1         On Error GoTo ErrHandler
          'Setting Application.Screenupdating to False works well except when we use _
           SendKeys in which case we just have to live with the screen flicker.
2         Application.ScreenUpdating = False

          'When resuming after use of SendKeys, we guard against the user having switched _
           sheet in the second or two that they have control.
3         If ResumePoint = rsafResumeAfterFirstSendKeys Or ResumePoint = rsafResumeAfterSecondSendKeys Then
4             On Error Resume Next
5             TempString = vbNullString
6             TempString = m_CurrentRange.Parent.Name
7             On Error GoTo ErrHandler
8             If TempString = vbNullString Then Throw "Unexpected Error: Sheet or Range has been deleted"
9             If ActiveSheet Is Nothing Then Throw "Unexpected Error: No Active Sheet found"
10            If Not ActiveSheet Is m_CurrentRange.Parent Then m_CurrentRange.Parent.Activate
11        End If

          'See if we are being called from ResizeArrayFormulaPart2 or ResizeArrayFormulaPart3 and _
           resume code execution at the appropriate point.
12        If ResumePoint = rsafResumeAfterFirstSendKeys Then
13            Set TempCells = ActiveSheet.Range(m_TempCellsAddress)
14            If Not sEquals(TempCells.FormulaArray, m_CurrentSizingFormula) Then
15                Throw "Unexpected error encountered when calculating the correct size for the array formula."
16            End If
17            ActiveSheet.Range(m_ActiveCellAddress).Select
18            Application.EnableCancelKey = xlErrorHandler
19            GoTo ResumeAfterFirstSendKeys
20        ElseIf ResumePoint = rsafResumeAfterSecondSendKeys Then
21            Application.EnableCancelKey = xlErrorHandler
22            GoTo ResumeAfterSecondSendKeys
23        Else
              'Disable user interuption until the successful call to BackUpRange
24            Application.EnableCancelKey = xlDisabled
25        End If

          'Populate module-level variables - these variables could now be statics in this method
26        m_Fit = Fit
27        m_DoneBackup = False
28        m_TempCellsAddress = vbNullString
29        m_VisibleRangeAddress = ActiveWindow.VisibleRange.address        'What happens if the window is split etc?
30        If TypeName(Selection) <> "Range" Then
31            ResizeArrayFormula = rsafAborted
32            Exit Function
33        End If
34        Set m_originalSelection = Selection
35        If m_originalSelection.Areas.Count > 1 Then
36            ResizeArrayFormula = rsafAborted
37            Exit Function
38        End If
39        Set m_CurrentRange = Selection        'Need this since m_currentRange is accessed in the error handler
40        Set m_RangeToSelectAfterError = Nothing
41        m_ActiveCellAddress = ActiveCell.address
42        If Not ActiveCell.HasFormula Then
43            ResizeArrayFormula = rsafAborted
44            Exit Function
45        End If

46        If Not UnprotectAsk(ActiveSheet, m_Title) Then Exit Function

          'Don't let the user do something mad like trying to resize an array to an entire sheet.
47        If Not m_Fit Then
48            If Selection.Cells.CountLarge > MAX_CELLS_FOR_RESIZE Then
49                Throw "Cannot resize to a range with more than " & Format$(MAX_CELLS_FOR_RESIZE, "###,###") + " cells."
50            End If
51        End If

52        m_origScrollArea = ActiveSheet.ScrollArea
53        If m_origScrollArea <> vbNullString Then ActiveSheet.ScrollArea = vbNullString

54        If ActiveCell.HasArray Then
55            Set m_CurrentRange = ActiveCell.CurrentArray
56        Else
57            Set m_CurrentRange = ActiveCell
58        End If

59        m_CurrentFormula = ActiveCell.Formula
60        If ThisWorkbook.isAddin Then
61            m_CurrentSizingFormula = "=sSizeOf(" & Right$(m_CurrentFormula, Len(m_CurrentFormula) - 1) & ")"
62        Else
63            m_CurrentSizingFormula = "=" & ThisWorkbook.Name & "!sSizeOf(" & Right$(m_CurrentFormula, Len(m_CurrentFormula) - 1) & ")"
64        End If

65        If Len(m_CurrentFormula) > 2086 Then Throw "The formula is too long to resize. The maximum length of formula" & _
              " which can be resized is 2086."
66        If Len(m_CurrentSizingFormula) > 2086 And m_Fit Then Throw "The formula is too long to auto-resize." & _
              " The maximum length of formula which can be auto-resized is 2076."

67        StatusBarText = "Resizing formula: " & m_CurrentFormula
68        If Len(StatusBarText) > 250 Then StatusBarText = Left$(StatusBarText, 247) & "..."
          'Set Calculation to manual and the ReferenceStyle to A1, set the StatusBar text, _
           set EditDirectlyInCell to False. When the ExcelStateHandler is set to Nothing, these changes are reversed.
69        Set m_ExcelState = CreateExcelStateHandler(xlCalculationManual, xlA1, , StatusBarText, False)

70        If Not m_Fit Then
71            Set m_NewRange = m_originalSelection
72        Else
73            Set TempCells = FindBlankCells()
74            m_TempCellsAddress = TempCells.address
75            If Not SetFormulaArray(TempCells, m_CurrentSizingFormula) Then
76                SetFormulaArrayViaSendKeys TempCells, m_CurrentSizingFormula
                  'The line below restarts execution at ResumeAfterFirstSendKeys, after allowing _
                   Excel to receive the keys sent via SendKeys.

77                Application.OnTime Now + TimeValue("00:00:01"), ThisWorkbook.Name & "!ResizeArrayFormulaPart2"
78                ResizeArrayFormula = rsafNotFinishedAwaitingSendKeys
79                Exit Function
80            End If
ResumeAfterFirstSendKeys:

81            If InStr(CStr(TempCells.Cells(1, 1).Value), ",") = 0 Then        'This can be the case if SendKeys method used, _
                                                                                even with Calculation Automatic.
82                TempCells.Calculate
83            End If
84            If InStr(CStr(TempCells.Cells(1, 1).Value), ",") = 0 Then
                  'A possible cause of failure here is that evaluating the VBA method sSizeOf fails when _
                   its argument is an extremely large array so we try one last time, using the native Excel _
                   functions ROWS and COLUMNS. The disadvantage of this approach is the fact that the big _
                   array formula is calculated twice, versus once with the sSizeOf approach. We do not _
                   attempt to use SendKeys when setting these formulas. Very large arrays combined with _
                   very long formulas is too much to cope with.
85                If SetFormulaArray(TempCells, "=ROWS(" & Right$(m_CurrentFormula, Len(m_CurrentFormula) - 1) & ")") Then
86                    If Not IsError(TempCells.Cells(1).Value) Then
87                        NumRows = TempCells(1).Value
88                        If SetFormulaArray(TempCells, "=COLUMNS(" & Right$(m_CurrentFormula, Len(m_CurrentFormula) - 1) & ")") Then
89                            If Not IsError(TempCells.Cells(1).Value) Then
90                                NumCols = TempCells(1).Value
91                            End If
92                        End If
93                    End If
94                End If
95            Else
96                On Error Resume Next
97                NumRows = Split(TempCells.Cells(1, 1).Value, ",")(0)
98                NumCols = Split(TempCells.Cells(1, 1).Value, ",")(1)
99                On Error GoTo ErrHandler
100           End If
101           If NumRows = 0 Or NumCols = 0 Then
102               Throw "Could not calculate the correct size for the resized formula"
103           End If
104           TempCells.ClearContents
              'If the TempCells.ClearContents worked then we don't need the address in the error handler cleanup
105           m_TempCellsAddress = vbNullString
106           If m_originalSelection.row + NumRows - 1 > ActiveSheet.Rows.Count Then
107               Throw "There are not enough rows in the worksheet to resize the formula."
108           End If
109           If m_originalSelection.Column + NumCols - 1 > ActiveSheet.Columns.Count Then
110               Throw "There are not enough columns in the worksheet to resize the formula."
111           End If
112           Set m_NewRange = m_originalSelection.Cells(1, 1).Resize(NumRows, NumCols)
113       End If

          'Search for reasons why the resize will fail so that we can post an informative error message.

          '1)Check that no merged calls are contained in the new range.
114       If IsNull(m_NewRange.MergeCells) Then
115           Throw MergedCellsErrorMessage(m_NewRange)
116       End If
          '2)Check that the cells that are to be part of m_NewRange but not in current range do not _
           include any array formulae that extend outside m_NewRange

117       Set RangeToTest = IWC(m_NewRange, m_CurrentRange)
118       If Not RangeToTest Is Nothing Then
119           ThrowIfError TestRangeForIntersectingFormulas(RangeToTest, m_RangeToSelectAfterError)
120       End If

          '3) Check that cells are either all locked or all unlocked, but first try to avoid this problem by locking all cells _
           cells being locked is the default state in a new worksheet.
121       If IsNull(m_NewRange.Locked) Then
122           On Error Resume Next
123           m_NewRange.Locked = True
124           On Error GoTo ErrHandler
125       End If
126       If IsNull(m_NewRange.Locked) Then
127           Set m_RangeToSelectAfterError = m_NewRange
128           Throw LockedCellsErrorMessage(m_NewRange)
129       End If
          '4) Test for intersection with ListObjects, also known as Tables
130       ThrowIfError TestRangeForIntersectingListObjects(m_NewRange, m_RangeToSelectAfterError)

          'In auto-size mode when we detect that we will be overwriting existing content, _
           get the user to confirm that that's OK.
131       If m_Fit Then
132           If Not RangeToTest Is Nothing Then
133               If Application.WorksheetFunction.CountA(RangeToTest) > 0 Then
134                   m_NewRange.Select
135                   Application.ScreenUpdating = True
136                   If MsgBoxPlus("Overwrite these cells?", vbOKCancel + vbDefaultButton2 + vbQuestion, m_Title) = vbCancel Then
137                       Application.ScreenUpdating = False
138                       m_DoneBackup = False        'User abandoned the resize so we don't want to populate the Undo menu
139                       ResizeArrayFormula = rsafAborted
140                       GoTo CleanupAndExit
141                   End If
142                   Application.ScreenUpdating = False
143               End If
144           End If
145       End If

          'Figure out a good range to backup so that we can have an UnDo capability. To avoid errors in _
           BackUpRange or RestoreRange, we need the range to be backed up not to contain only part of _
           any array formula taking into account how the array formulas are positioned in both the _
           pre-resize state and the post-resize state.
          Dim RangeToBackup As Range
          Dim TempRange As Range
146       On Error Resume Next
          'We are only populating the temporary cells to give .CurrentRegion the behaviour we want _
           when calculating the RangeToBackUp so we only need to populate the non-blank cells on the perimeter.
147       Set TempRange = BlankCellsInRange(PerimeterCells(m_NewRange))

148       On Error GoTo ErrHandler
149       If Not TempRange Is Nothing Then
150           Application.EnableCancelKey = xlDisabled
151           TempRange.Value = "Foo"
152       End If
          'The .CurrentRegion is prone to throwing errors - e.g. if the sheet's ScrollArea property is _
           small but throwing an error on that line would be bad - those pesky "Foo"s are still there. _
           So clean them up before testing if the call to .CurrentRegion worked OK.
153       On Error Resume Next
154       Set RangeToBackup = Application.Union(m_CurrentRange, m_NewRange).CurrentRegion
155       On Error GoTo ErrHandler

156       If Not TempRange Is Nothing Then
157           TempRange.ClearContents
158       End If

159       If RangeToBackup Is Nothing Then Throw "Unexpected Error: Failed to get CurrentRegion of range"

160       If RangeToBackup.Cells.CountLarge > MAX_CELLS_FOR_UNDO Then
161           CleanUpUndoBuffer shUndo
162           Prompt = "Undo (Ctrl+Z) is not available for an operation as large as this." + vbLf + vbLf + _
                  "Continue without Undo?"
163           If MsgBoxPlus(Prompt, vbYesNo + vbDefaultButton2 + vbQuestion, m_Title) <> vbYes Then
164               GoTo CleanupAndExit
165           End If
166       Else
167           BackUpRange RangeToBackup, shUndo, Application.Union(m_CurrentRange, m_NewRange)
168           m_DoneBackup = True
169       End If

          'Now that we've done the backup (or the user has accepted that it won't be done) we can enable user interrupt.
170       Application.EnableCancelKey = xlErrorHandler

171       m_CurrentRange.ClearContents
172       If SetFormulaArray(m_NewRange, m_CurrentFormula) Then
173           m_NewRange.Select
174       Else
              'It's possible that SetFormulaArray populated m_NewRange with a formula but the _
               formula was incorrect, so clear the range.
175           m_NewRange.ClearContents
176           SetFormulaArrayViaSendKeys m_NewRange, m_CurrentFormula
              'The line below restarts execution at ResumeAfterSecondSendKeys, after allowing _
               Excel to receive the keys sent via SendKeys.
177           Application.OnTime Now + TimeValue("00:00:01"), ThisWorkbook.Name & "!ResizeArrayFormulaPart3"
178           ResizeArrayFormula = rsafNotFinishedAwaitingSendKeys
179           Exit Function
180       End If

ResumeAfterSecondSendKeys:

          'Final check that we have indeed correctly set the formula!
181       If Not sEquals(m_NewRange.FormulaArray, m_CurrentFormula) Then
182           Debug.Print String(100, "=")
183           Debug.Print "Assertion Failure in method " + ThisWorkbook.Name + "!ResizeArrayFormula"
184           Debug.Print "Original formula and Resized formula differed, so resize attempt was aborted."
185           Debug.Print "Original Formula:"
186           Debug.Print m_CurrentFormula
187           Debug.Print "Resized Formula:"
188           Debug.Print m_NewRange.FormulaArray
189           Debug.Print String(100, "=")
190           Throw "Unknown error. The attempt to resize the formula resulted in " & _
                  "unintended changes to the formula and so the attempt was aborted."
191       End If

          'Handle the view port, but ignore errors as unimportant.
192       On Error Resume Next
193       Application.Goto ActiveSheet.Range(m_VisibleRangeAddress)
194       m_NewRange.Select
195       On Error GoTo ErrHandler

196       ResizeArrayFormula = rsafWorkedCorrectly
CleanupAndExit:
          'Revert Excel's state and sheet ScrollArea
197       Set m_ExcelState = Nothing
198       If ActiveSheet.ScrollArea <> m_origScrollArea Then ActiveSheet.ScrollArea = m_origScrollArea
199       If m_DoneBackup Then
              'Set up Undo text. The check below that the UndoBuffer sheet exists ought to be redundant, _
               but there's no harm in having it.
200           If IsUndoAvailable(shUndo) Then
                  Dim UndoText As String
201               If Len(m_CurrentFormula) > 53 Then
202                   UndoText = "'" & Left$(m_CurrentFormula, 50) & "..." & "'"
203               Else
204                   UndoText = "'" & m_CurrentFormula & "'"
205               End If

206               UndoText = "Undo " & IIf(m_Fit, "Auto ", vbNullString) & "Resize Array Formula " & UndoText & " to " & _
                      AddressND(Selection)

207               Application.OnUndo UndoText, "RestoreRange"
208           End If
209       End If

210       Exit Function
ErrHandler:
211       Prompt = "Something went wrong:" + vbLf + vbLf + Err.Description
212       Set m_ExcelState = Nothing        'Revert Excel's state

213       On Error Resume Next
214       TempString = vbNullString
215       TempString = m_CurrentRange.Parent.Name        'Check that the sheet still exists - is it possible for _
                                                          the user to have deleted it while SendKeys was processing?
216       If TempString <> vbNullString Then
217           If m_DoneBackup Then
218               RestoreRange
219               ResizeArrayFormula = rsafFailedButCorrectCleanup
220           Else
221               ResizeArrayFormula = rsafAborted
222           End If
223           With m_CurrentRange.Parent
224               If m_TempCellsAddress <> vbNullString Then .Range(m_TempCellsAddress).ClearContents
225               Application.Goto .Range(m_VisibleRangeAddress)
226           End With
227           If Not m_RangeToSelectAfterError Is Nothing Then
228               m_RangeToSelectAfterError.Select
229           Else
230               m_CurrentRange.Select
231           End If
232       Else
233           ResizeArrayFormula = rsafFailed
234       End If
          'Revert Scroll area
235       If ActiveSheet.ScrollArea <> m_origScrollArea Then ActiveSheet.ScrollArea = m_origScrollArea
236       Application.ScreenUpdating = True
237       MsgBoxPlus Prompt, vbExclamation, m_Title
238       Exit Function
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LockedCellsErrorMessage
' Author    : Philip Swannell
' Date      : 17-May-2015
' Purpose   : Construction of a helpful error message when locked/unlocked cells get in the way.
' -----------------------------------------------------------------------------------------------------------------------
Private Function LockedCellsErrorMessage(TheRange As Range)
          Dim c As Range
          Dim FoundLocked As Boolean
          Dim FoundUnlocked As Boolean
          Dim LockedCell As Range
          Dim LockedCellAddress As String
          Dim UnlockedCell As Range
          Dim UnlockedCellAddress As String

1         On Error GoTo ErrHandler
2         For Each c In TheRange.Cells
3             If Not FoundLocked And c.Locked Then
4                 FoundLocked = True
5                 LockedCellAddress = AddressND(c)
6                 Set LockedCell = c
7             End If
8             If Not FoundUnlocked And c.Locked = False Then
9                 FoundUnlocked = True
10                UnlockedCellAddress = AddressND(c)
11                Set UnlockedCell = c
12            End If
13            If FoundLocked And FoundUnlocked Then Exit For
14        Next
15        LockedCellsErrorMessage = "Cannot enter an array formula into a range of cells that are not all locked or all unlocked." + _
              " For example, cell " + LockedCellAddress + " is locked but cell " + UnlockedCellAddress + " is unlocked."
16        Exit Function
ErrHandler:
17        LockedCellsErrorMessage = "Cannot enter an array formula into a range of cells that are not all locked or all unlocked."
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MergedCellsErrorMessage
' Author    : Philip Swannell
' Date      : 28-Oct-2013
' Purpose   : Construction of a helpful error message when merged cells get in the way
' -----------------------------------------------------------------------------------------------------------------------
Private Function MergedCellsErrorMessage(TheRange As Range)
          Dim c As Range
1         On Error GoTo ErrHandler
2         For Each c In TheRange.Cells
3             If c.MergeCells Then
4                 MergedCellsErrorMessage = "Array formulas are not valid in merged cells such as cell " & _
                      AddressND(c) & "."
5                 Exit Function
6             End If
7         Next
8         Exit Function
ErrHandler:
9         MergedCellsErrorMessage = "Array formulas are not valid in merged cells"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeArrayFormulaPart2
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : In the event that we have to resort to SendKeys for setting the "sizing formula", this
'             method is called via Application.OnTime and calls back into ResizeArrayFormula,
'             setting the appropriate point to resume code execution.
' -----------------------------------------------------------------------------------------------------------------------
Sub ResizeArrayFormulaPart2()
1         ResizeArrayFormula m_Fit, rsafResumeAfterFirstSendKeys
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeArrayFormulaPart3
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : In the event that we have to resort to SendKeys for setting the formula, this
'             method is called via Application.OnTime and calls back into ResizeArrayFormula,
'             setting the appropriate point to resume code execution.
' -----------------------------------------------------------------------------------------------------------------------
Sub ResizeArrayFormulaPart3()
1         ResizeArrayFormula m_Fit, rsafResumeAfterSecondSendKeys
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestRangeForIntersectingFormulas
' Author    : Philip Swannell
' Date      : 25-Oct-2013
' Purpose   : In the event that there exists an array formula which is partly inside RangeToTest
'             and partly outside it, then this method throws an error. RangeToTest may be a
'             multiple area range.
' -----------------------------------------------------------------------------------------------------------------------
Private Function TestRangeForIntersectingFormulas(RangeToTest As Range, ByRef RangeToSelect)
          Dim BlankCellsInRangeToTest As Range
          Dim c As Range
          Dim FormulaCellsInRangeToTest As Range
          Dim nonEmptiesFound As Boolean

1         On Error GoTo ErrHandler
2         Set BlankCellsInRangeToTest = BlankCellsInRange(RangeToTest)
3         If BlankCellsInRangeToTest Is Nothing Then
4             nonEmptiesFound = True
5         Else
6             If Not RangesIdentical(BlankCellsInRangeToTest, RangeToTest) Then
7                 nonEmptiesFound = True
8             End If
9         End If
10        If nonEmptiesFound Then
11            Set FormulaCellsInRangeToTest = CellsWithFormulasInRange(RangeToTest)
12            If Not FormulaCellsInRangeToTest Is Nothing Then

13                For Each c In FormulaCellsInRangeToTest.Cells
14                    If c.HasArray Then
15                        If Application.Intersect(c.CurrentArray, m_NewRange).Cells.CountLarge <> c.CurrentArray.Cells.CountLarge Then
16                            Set RangeToSelect = Application.Union(m_NewRange, c.CurrentArray)
17                            TestRangeForIntersectingFormulas = "#You cannot change part of the array at " & AddressND(c.CurrentArray) & "!"
18                            Exit Function
19                        End If
20                    End If
21                Next c
22            End If
23        End If
24        Exit Function
ErrHandler:
25        TestRangeForIntersectingFormulas = "#TestRangeForIntersectingFormulas (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestRangeForIntersectingListObjects
' Author    : Philip Swannell
' Date      : 17-Jul-2015
' Purpose   : The cells of ListObjects cannot contain multi-cell array formulas
' -----------------------------------------------------------------------------------------------------------------------
Private Function TestRangeForIntersectingListObjects(RangeToTest As Range, ByRef RangeToSelect As Range)
          Dim lo As ListObject
          Dim LORange As Range

1         On Error GoTo ErrHandler
2         If RangeToTest.Cells.CountLarge > 1 Then
3             For Each lo In RangeToTest.Parent.ListObjects
4                 Set LORange = ListObjectRange(lo)
5                 If Not LORange Is Nothing Then
6                     If Not Application.Intersect(LORange, RangeToTest) Is Nothing Then
7                         If Not IntersectWithComplement(LORange, RangeToTest) Is Nothing Then
8                             Set RangeToSelect = RangeToTest
9                             TestRangeForIntersectingListObjects = "#Multi-cell array formulas are not allowed in tables, such as " & lo.Name & " at " & AddressND(LORange) & "!"
10                            Exit Function
11                        End If
12                    End If
13                End If
14            Next
15        End If

16        Exit Function
ErrHandler:
17        TestRangeForIntersectingListObjects = "#TestRangeForIntersectingListObjects (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ListObjectRange
' Author    : Philip Swannell
' Date      : 17-Jul-2015
' Purpose   : Returns the Range that a ListObject (aka Table) occupies
' -----------------------------------------------------------------------------------------------------------------------
Private Function ListObjectRange(lo As ListObject, Optional ExpandBy1RowAndCol As Boolean) As Range
          Dim DataBodyRange As Range
          Dim HeaderRange As Range
          Dim Result As Range
          Dim TotalsRange As Range

1         On Error GoTo ErrHandler
2         On Error Resume Next
3         Set DataBodyRange = lo.DataBodyRange
4         Set HeaderRange = lo.HeaderRowRange
5         Set TotalsRange = lo.TotalsRowRange
6         On Error GoTo ErrHandler
7         Set Result = UnionOfRanges(DataBodyRange, HeaderRange, TotalsRange)
8         If ExpandBy1RowAndCol Then
9             If Result.row + Result.Rows.Count - 1 < Result.Parent.Rows.Count Then
10                Set Result = Result.Resize(Result.Rows.Count + 1)
11            End If
12            If Result.Column + Result.Columns.Count - 1 < Result.Parent.Columns.Count Then
13                Set Result = Result.Resize(, Result.Columns.Count + 1)
14            End If
15        End If
16        Set ListObjectRange = Result

17        Exit Function
ErrHandler:
18        Set ListObjectRange = Nothing
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetFormulaArray
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Wrapper to .FormulaArray = so that we can try the SendKeys approach if the
'             method fails, or sets the formula to the wrong formula - see remarks below.
' -----------------------------------------------------------------------------------------------------------------------
Function SetFormulaArray(TheRange As Range, TheFormula As String) As Boolean
          'According to the documentation (http://msdn.microsoft.com/en-us/library/office/ff837104.aspx)
          'the FormulaArray method requires that the formula passed in is in R1C1 notation. However, if
          'you pass in a formula in A1 notation it will generally work. But that's bad because some
          'formulae can be syntactically correct in both notations but have different meaning, e.g. =Sum(C1)
          'means the sum of all cells in the left column in R1C1 notation or it means the sum of the contents
          'if one cell (third col, first row) in A1 notation. So for safety after setting the formula we check
          'that it indeed became what we intended it to become. In the unhappy case that the formula is not
          'what we want it to be then we try passing in the formula in A1 notation. If that second attempt fails
          'this method returns False and the calling method can then try the alternative method SetFormulaArrayViaSendKeys

          Dim ConvertedFormula As String

1         On Error GoTo ErrHandler1
          'Line below throws an error if the formula is long (>255)
2         ConvertedFormula = Application.ConvertFormula(TheFormula, xlA1, xlR1C1, , TheRange.Cells(1, 1))

          'Line below would throw an error if there were merged cells or some locked, some unlocked, but we have guarded against that in the calling method.
          'Nightmare possibility: The formula being processed is written in VBA and the VBA throws an error when the formula is entered. I don't think we can get
          'this code to behave gracefully in that circumstance.
3         TheRange.FormulaArray = ConvertedFormula

4         On Error GoTo ErrHandler2
          'Have we set the formula correctly?
5         If LCase$(TheRange.FormulaArray) = LCase$(TheFormula) Then
6             SetFormulaArray = True
7             Exit Function
8         Else
              'The formula has gone in wrong, try again using A1 notation rather than R1C1 notation.
9             TheRange.FormulaArray = TheFormula
10            If LCase$(TheRange.FormulaArray) = LCase$(TheFormula) Then
11                SetFormulaArray = True
12                Exit Function
13            End If
14            SetFormulaArray = False
15        End If
16        Exit Function
ErrHandler1:
17        SetFormulaArray = False
18        Exit Function
ErrHandler2:
19        Throw "#SetFormulaArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeGoodForArrayFormula
' Author    : Philip Swannell
' Date      : 17-May-2015
' Purpose   : Encapsulate whether a range is empty and can accept an array formula.
'             Would be slow for large ranges, since it loops over all cells in range
'             but we only use it for a small range.
' -----------------------------------------------------------------------------------------------------------------------
Private Function RangeGoodForArrayFormula(TheRange As Range) As Boolean
1         On Error GoTo ErrHandler
          Dim c As Range
          Dim t As ListObject

2         For Each c In TheRange.Cells
3             If Not IsEmpty(c.Value) Then
4                 RangeGoodForArrayFormula = False
5                 Exit Function
6             End If
7         Next c

8         If IsNull(TheRange.Locked) Then
9             RangeGoodForArrayFormula = False
10        ElseIf IsNull(TheRange.MergeCells) Then
11            RangeGoodForArrayFormula = False
12        ElseIf TheRange.MergeCells Then
13            RangeGoodForArrayFormula = False
14        Else
              'Can't have array formulas in Tables, aka ListObjects, also putting data just to the right
              ' or just below a Table can cause the table to automatically expand
15            For Each t In TheRange.Parent.ListObjects
16                If Not Application.Intersect(TheRange, ListObjectRange(t, True)) Is Nothing Then
17                    RangeGoodForArrayFormula = False
18                    Exit Function
19                End If
20            Next
21            RangeGoodForArrayFormula = True
22        End If

23        Exit Function
ErrHandler:
24        RangeGoodForArrayFormula = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FindBlankCells
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : When doing auto-sizing we need two "spare" cells on the active sheet into which
'             enter the "sizing formula" =sSizeOf(<TheExistingFormula>). If possible we want
'             these cells to be in the currently visible range because this reduces screen
'             flicker in the event that  we cannot switch off screen updating because we are
'             having to use SendKeys. Also the cells found must not be a merged cells - array
'             formulae cannot be entered in such cells.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FindBlankCells() As Range
          Dim BlankCells As Range
          Dim c As Range
          Dim i As Long
          Dim NumRows As Long
          Dim TempRange As Range
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         For i = -3 To 10
3             If i = -3 Then
                  'Search in cells to the right of active cell and in the visible range
4                 Set TempRange = Application.Intersect(ActiveCell.Resize(1, ActiveSheet.Columns.Count - ActiveCell.Column + 1), _
                      ActiveWindow.VisibleRange)
5             ElseIf i = -2 Then
                  'Search in the current row and in the visible range
6                 Set TempRange = Application.Intersect(ActiveCell.EntireRow, ActiveWindow.VisibleRange)
7             ElseIf i = -1 Then
                  'Search in the visible range
8                 Set TempRange = ActiveWindow.VisibleRange
9             Else
                  'On the remaining passes search in blocks of cells below right of the UsedRange
10                Set TempRange = Nothing
11                On Error Resume Next
12                With ActiveSheet.UsedRange
13                    Set TempRange = .Cells(.Rows.Count + 1 + i * 10, .Columns.Count + 1 + i * 10).Resize(10, 10)
14                    On Error GoTo ErrHandler
15                End With
16            End If
17            If Not TempRange Is Nothing Then
18                Set BlankCells = BlankCellsInRange(TempRange)
19                If Not BlankCells Is Nothing Then
20                    For Each c In BlankCells.Cells
21                        If RangeGoodForArrayFormula(c.Resize(1, 2)) Then
22                            Set FindBlankCells = c.Resize(1, 2)
23                            Exit Function
24                        ElseIf RangeGoodForArrayFormula(c.Resize(2, 1)) Then
25                            Set FindBlankCells = c.Resize(2, 1)
26                            Exit Function
27                        End If
28                    Next c
29                End If
30            End If
31            Set BlankCells = Nothing
32        Next i

          'Otherwise look for a pair of unmerged blank cells anywhere on the sheet in each column _
           of the sheet, try the cells one and two below the lowest non-blank cell
33        Set ws = ActiveSheet
34        NumRows = ws.Rows.Count
35        For i = 1 To ws.Columns.Count
36            Set c = ws.Cells(NumRows, i).End(xlUp).Cells(2, 1)
37            If c.row < NumRows - 1 Then
38                If RangeGoodForArrayFormula(c.Resize(2, 1)) Then
39                    Set FindBlankCells = c.Resize(2, 1)
40                    Exit Function
41                End If
42            End If
43        Next i
44        Throw "Cannot find pair of blank cells for temporary use"
45        Exit Function
ErrHandler:
46        Throw "#FindBlankCells (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSizeOf
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Utility function used by ResizeArrayFormula as part of the "sizing formula"
' -----------------------------------------------------------------------------------------------------------------------
Function sSizeOf(TheArray As Variant) As Variant
Attribute sSizeOf.VB_Description = "Returns a string describing the size of TheArray. The Return is the number of rows followed by the number of columns with a comma between them. The function is required by the automatic array resizing method (SOLUM > Resize Array Formula > Auto Resize)."
Attribute sSizeOf.VB_ProcData.VB_Invoke_Func = " \n24"

1         On Error GoTo ErrHandler
2         If TypeName(TheArray) = "Range" Then
3             sSizeOf = CStr(TheArray.Rows.Count) + "," + CStr(TheArray.Columns.Count)
4         Else
5             sSizeOf = CStr(sNRows(TheArray)) & "," & CStr(sNCols(TheArray))
6         End If
7         Exit Function
ErrHandler:
8         Throw "#sSizeOf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetFormulaArrayViaSendKeys
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Set the formula array using SendKeys. Rather than sending the entire formula
'             we set the top left cell contents to be the formula escaped with an apostrophe
'             and then send keys to delete the apostrophe and do the Ctrl+Shift+Enter.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SetFormulaArrayViaSendKeys(TheRange As Range, TheFormula As String)

          Dim i As Long
          Dim origHeights As Variant
          Dim origHorizontalAlignment As Long

1         On Error GoTo ErrHandler
2         TheRange.Select

          'If TheFormula contains carriage return characters then entering it as a string will likely change the height of the cell so fix up
3         origHeights = sReshape(0, TheRange.Rows.Count, 1)
4         For i = 1 To TheRange.Rows.Count
5             origHeights(i, 1) = TheRange.Cells(i, 1).RowHeight
6         Next i

7         With TheRange.Cells(1, 1)
8             origHorizontalAlignment = .HorizontalAlignment
9             .Value = "'" & TheFormula
10            .HorizontalAlignment = origHorizontalAlignment
11            For i = 1 To TheRange.Rows.Count
12                TheRange.Cells(i, 1).RowHeight = origHeights(i, 1)
13            Next i
14        End With

15        AppActivate Application.caption, True
          'Application.SendKeys "{F2}^{HOME}{DELETE}^+~", True
          'Avoid sending the {DELETE} character since if this method is assigned to _
           Ctrl+Alt+aCharacter then Windows can end up intercepting Ctrl+Alt+Delete _
           i.e. lock the PC! So instead of {DELETE} we use {BACKSPACE}
16        Application.SendKeys "{F2}^{HOME}{RIGHT}{BACKSPACE}^+~", True

17        Exit Sub
ErrHandler:
18        Throw "#SetFormulaArrayViaSendKeys (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnionOfRanges
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Union of multiple ranges some of which may be Nothing
' -----------------------------------------------------------------------------------------------------------------------
Function UnionOfRanges(ParamArray Ranges()) As Range
          Dim i As Long
          Dim Result As Range

1         On Error GoTo ErrHandler

2         For i = LBound(Ranges, 1) To UBound(Ranges, 1)
3             If Not Ranges(i) Is Nothing Then
4                 If TypeName(Ranges(i)) <> "Range" Then Throw "Parameters must all be Nothing or Range objects"
5                 If Result Is Nothing Then
6                     Set Result = Ranges(i)
7                 Else
8                     Set Result = Application.Union(Result, Ranges(i))
9                 End If
10            End If
11        Next

12        If Not Result Is Nothing Then
13            Set UnionOfRanges = Result
14        End If

15        Exit Function
ErrHandler:
16        Throw "#UnionOfRanges (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangesIdentical
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Tests whether two multi-area ranges are "really" the same i.e. a cell is
'             in one range if and only if it is in the other. Should give True for
'             Range("$O$9:$AG$28,$N$10:$AG$28") and Range("$N$10:$N$28,$O$9:$AG$28")
' -----------------------------------------------------------------------------------------------------------------------
Function RangesIdentical(RangeA As Range, RangeB As Range) As Boolean

1         On Error GoTo ErrHandler
2         If RangeA.Parent Is RangeB.Parent Then
3             If IntersectWithComplement(RangeA, RangeB) Is Nothing Then
4                 If IntersectWithComplement(RangeB, RangeA) Is Nothing Then
5                     RangesIdentical = True
6                 End If
7             End If
8         End If
9         Exit Function
ErrHandler:
10        Throw "#RangesIdentical (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : testRangesIdentical
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Quick test harness...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestRangesIdentical()
          Dim i As Long
          Dim RangeA As Range
          Dim RangeB As Range
1         Application.Workbooks.Add

2         For i = 1 To 4
3             Select Case i
                  Case 1
4                     Set RangeA = ActiveSheet.Range("$A$1:$A$10,$B$1:$B$10,$C$1:$C$10,$D$1:$D$10,$E$1:$E$10,$F$1:$F$10")
5                     Set RangeB = ActiveSheet.Range("A1:F10")
6                 Case 2
7                     Set RangeA = ActiveSheet.Range("$A$1:$A$10,$B$1:$B$10,$C$1:$C$10,$D$1:$D$10,$E$1:$E$10,$F$1:$F$10")
8                     Set RangeB = ActiveSheet.Range("$A$1:$F$1,$A$2:$F$2,$A$3:$F$3,$A$4:$F$4,$A$5:$F$5,$A$6:$F$6,$A$7:$F$7,$A$8:$F$8,$A$9:$F$9,$A$10:$F$10")
9                 Case 3
10                    Set RangeA = ActiveSheet.Range("$A$1:$A$10,$B$1:$B$10,$C$1:$C$10,$D$1:$D$10,$E$1:$E$10,$F$1:$F$10")
11                    Set RangeB = ActiveSheet.Range("$A$1:$F$1,$A$2:$F$2,$A$3:$F$3,$A$4:$F$4,$A$5:$F$5,$A$6:$F$6,$A$7:$F$7,$A$8:$F$8,$A$9:$F$9,$A$10:$F$10,G10")
12                Case 4
13                    Set RangeA = ActiveSheet.Range("$A$1:$A$10,$B$2:$B$10,$C$1:$C$10,$D$1:$D$10,$E$1:$E$10,$F$1:$F$10")
14                    Set RangeB = ActiveSheet.Range("$A$1:$F$1,$A$2:$F$2,$A$3:$F$3,$A$4:$F$4,$A$5:$F$5,$A$6:$F$6,$A$7:$F$7,$A$8:$F$8,$A$9:$F$9,$A$10:$F$10")
15            End Select
16            ActiveSheet.UsedRange.Clear
17            RangeA.Value = "RangeA"
18            RangeB.Interior.ColorIndex = 15
19            MsgBoxPlus RangesIdentical(RangeA, RangeB)
20        Next
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PerimeterCells
' Author    : Philip Swannell
' Date      : 31-Oct-2013
' Purpose   : For an input Range, returns the range which is the top plus bottom rows and
'             left and right columns.
' -----------------------------------------------------------------------------------------------------------------------
Private Function PerimeterCells(TheRange As Range)
          Dim a As Range
1         On Error GoTo ErrHandler
2         Set PerimeterCells = TheRange.Areas(1).Cells(1, 1)

3         For Each a In TheRange.Areas
4             With a
5                 If .Rows.Count <= 2 Or .Columns.Count <= 2 Then
6                     Set PerimeterCells = Application.Union(PerimeterCells, a)
7                 Else
8                     Set PerimeterCells = Application.Union(PerimeterCells, .Rows(1), _
                          .Rows(.Rows.Count), .Columns(1), _
                          .Columns(.Columns.Count))
9                 End If
10            End With
11        Next a

12        Exit Function
ErrHandler:
13        Throw "#PerimeterCells (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BlankCellsInRange
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Need this method to workaround some gotchas in .SpecialCells(xlCellTypeBlanks):
'             a) If you pass in a single cell that's equivalent to passing in the entire
'             UsedRange of the parent sheet :-(
'             b) If you pass in a Range that goes outside the UsedRange then the return only
'             mentions the blank cells inside the UsedRange :-(
'             This method gets around these gotchas so that it correctly does what it says on the tin.
' -----------------------------------------------------------------------------------------------------------------------
Function BlankCellsInRange(TheRange As Range) As Range
          Dim BlankCellsInIntersection As Range
          Dim IntersectionWithUsedRange As Range

1         On Error GoTo ErrHandler

2         If RangeHasJustOneCell(TheRange) Then        '.SpecialCells no use in this case - if we pass in a _
                                                        single cell that's equivalent to passing in the entire UsedRange
3             If IsEmpty(TheRange.Value) Then
4                 Set BlankCellsInRange = TheRange
5             End If
6             Exit Function
7         End If

8         Set IntersectionWithUsedRange = Application.Intersect(TheRange, TheRange.Parent.UsedRange)
9         If IntersectionWithUsedRange Is Nothing Then
10            Set BlankCellsInRange = TheRange
11            Exit Function
12        End If

13        If RangeHasJustOneCell(IntersectionWithUsedRange) Then
14            If IsEmpty(IntersectionWithUsedRange.Value) Then
15                Set BlankCellsInIntersection = IntersectionWithUsedRange
16            End If
17        Else
18            On Error Resume Next
19            Set BlankCellsInIntersection = TheRange.SpecialCells(xlCellTypeBlanks)
20            On Error GoTo ErrHandler
21        End If

          'Method UnionOfRanges handles Nothing gracefully, and method IntersectWithComplement may have a return of type Nothing
22        Set BlankCellsInRange = UnionOfRanges(BlankCellsInIntersection, IntersectWithComplement(TheRange, TheRange.Parent.UsedRange))

23        Exit Function
ErrHandler:
24        Throw "#BlankCellsInRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CellsWithFormulasInRange
' Author    : Philip Swannell
' Date      : 21-Oct-2013
' Purpose   : Returns a Range consisting of all the cells in TheRange that contain formulas.
'             Returns Nothing if there are no such cells. Need this method to work around two
'             gotchas in .SpecialCells: 1) doesn't work for a single cell. 2) raises an error
'             if no cells contain formulae.
' -----------------------------------------------------------------------------------------------------------------------
Function CellsWithFormulasInRange(TheRange As Range) As Range
1         If RangeHasJustOneCell(TheRange) Then
2             If TheRange.HasFormula Then
3                 Set CellsWithFormulasInRange = TheRange
4             End If
5             Exit Function
6         End If
7         On Error Resume Next
8         Set CellsWithFormulasInRange = TheRange.SpecialCells(xlCellTypeFormulas)
9         On Error GoTo 0
End Function

Function NonBlankCellsInRange(TheRange As Range)
1         If RangeHasJustOneCell(TheRange) Then
2             If Not (IsEmpty(TheRange.Value)) Then
3                 Set NonBlankCellsInRange = TheRange
4             End If
5             Exit Function
6         End If
7         Set NonBlankCellsInRange = TheRange.SpecialCells(xlCellTypeConstants)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeHasJustOneCell
' Author    : Philip Swannell
' Date      : 28-Oct-2013
' Purpose   : Work around the problem that executing TheRange.Cells.Count can give overflow error
' -----------------------------------------------------------------------------------------------------------------------
Function RangeHasJustOneCell(TheRange As Range) As Boolean
1         On Error GoTo ErrHandler

2         RangeHasJustOneCell = TheRange.Cells.CountLarge = 1

3         Exit Function
ErrHandler:
4         RangeHasJustOneCell = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FitArrayFormula2
' Author     : Philip Swannell
' Date       : 01-Aug-2019
' Purpose    : Use this instead of FitArrayFormula in dynamic-array-aware Excel
' -----------------------------------------------------------------------------------------------------------------------
Sub FitArrayFormula2()
          Dim CurrentArrayRange As Range
          Dim TheFormula As String
          Dim TheSize
          Dim NR As Long, NC As Long
          Dim TargetRange As Range
          Dim TheCell As Range

1         On Error GoTo ErrHandler
2         If Not ActiveCell Is Nothing Then
3             Set TheCell = ActiveCell
4             If CStr(TheCell.Value) = "Error 2045" Then ' "#SPILL error, always in only one cell.
5                 If Not UnprotectAsk(ActiveSheet, m_Title) Then Exit Sub
6                 TheFormula = TheCell.Formula
7                 TheCell.Formula2 = TheFormula
8                 If CStr(TheCell.Value) = "Error 2045" Then 'still getting #SPILL error
9                     TheCell.Formula2 = "=sSizeOf(" & Mid(TheFormula, 2) & ")"
10                    TheSize = TheCell.Value
11                    TheCell.Formula2 = TheFormula
12                    If VarType(TheSize) = vbString Then
13                        NR = CLng(sStringBetweenStrings(TheSize, , ","))
14                        NC = CLng(sStringBetweenStrings(TheSize, ","))
15                        If TheCell.row + NR - 1 > ActiveSheet.Rows.Count Or TheCell.Column + NC - 1 > ActiveSheet.Columns.Count Then
16                            Throw "Formula spills beyond the edge of the worksheet", True
17                        End If
18                        Set TargetRange = TheCell.Resize(NR, NC)
19                        If NumBlanksInRange(TargetRange) < TargetRange.Cells.CountLarge - 1 Then
20                            Application.ScreenUpdating = True
21                            TargetRange.Select
22                            If MsgBoxPlus("Overwrite these cells?", vbOKCancel + vbDefaultButton2 + vbQuestion, m_Title) = vbOK Then
23                                TargetRange.ClearContents
24                                TheCell.Formula2 = TheFormula
25                            End If
26                        End If
27                    End If
28                End If
29            ElseIf TheCell.HasArray Then
30                If Not UnprotectAsk(ActiveSheet, m_Title) Then Exit Sub
31                Set CurrentArrayRange = TheCell.CurrentArray
32                TheFormula = CurrentArrayRange.Cells(1, 1).Formula
33                CurrentArrayRange.ClearContents
34                TheCell.Formula2 = TheFormula
35            End If
36        End If
37        Exit Sub
ErrHandler:
38        Throw "#FitArrayFormula2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
