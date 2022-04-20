VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
1         AssignKeys False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
1         CleanUpUndoBuffer shUndo
2         CleanUpUndoBuffer shUndo2
End Sub

Private Sub Workbook_Open()
1         On Error GoTo ErrHandler
          Dim Res As Variant
          Dim wb As Excel.Workbook
2         Application.ScreenUpdating = False

3         Set wb = Application.Workbooks.Add

          Dim R As Range
          'Line below "seeds" the Find and Replace dialog with "Look in" set to Values, which is what I want 99% of the time
4         Set R = Cells.Find(what:="", after:=ActiveCell, LookIn:=xlValues, LookAt:= _
              xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
              , SearchFormat:=False)

5         AssignKeys True
6         EnsureAppObjectExists

          'See also method SolumAddinSettings that allows the user to control the contents of this Registry key
7         Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "Application.Calculation", "Nothing"))
              Case LCase$("xlCalculationAutomatic")
                  'Cannot change Application.Calculation if no workbooks are open
8                 Application.Calculation = xlCalculationAutomatic
9                 Application.CalculateBeforeSave = True
10            Case LCase$("xlCalculationSemiautomatic")
11                Application.Calculation = xlCalculationSemiautomatic
12                Application.CalculateBeforeSave = True
13            Case LCase$("xlCalculationManual")
14                Application.Calculation = xlCalculationManual
15                Application.CalculateBeforeSave = False
16        End Select

17        If Not wb Is Nothing Then
18            wb.Windows(1).Visible = False
19            wb.Saved = True
20        End If

          'If BERT is installed then check that required packages are installed and install them if not.
21        On Error Resume Next
22        Res = Application.Run("BERT.Exec", "1+1")
23        On Error GoTo ErrHandler
24        If sEquals(Res, 2) Then
25            CheckR vbNullString, gPackages, vbNullString, vbNullString, False
26        End If
27        Application.OnTime Now(), ThisWorkbook.Name & "!RecordInstalledFonts" 'RecordInstalledFonts does not work if called directly
28        Application.OnRepeat vbNullString, vbNullString 'otherwise Repeat button shows "Repeat Macros (Ctrl Y)" and, if clicked, adds a new workbook
29        SetApplicationCaptions
30        Exit Sub
ErrHandler:
31        SomethingWentWrong "#Workbook_Open (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

