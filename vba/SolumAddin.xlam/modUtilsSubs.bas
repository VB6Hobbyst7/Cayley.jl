Attribute VB_Name = "modUtilsSubs"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PeekCell
' Author     : Philip Swannell
' Date       : 23-Apr-2020
' Purpose    : For when you can't see the entire contents of a cell - e.g. it has a long error message. Attached to F12
' -----------------------------------------------------------------------------------------------------------------------
Sub PeekCell()
          Dim Prompt
          Dim Title
          Dim c As Range
          Dim D As Range
          Dim TheValue As String
          Dim TheFormula As String
          Dim Res As VbMsgBoxResult

1         On Error GoTo ErrHandler
2         If Not ActiveCell Is Nothing Then
3             Set c = ActiveCell
4             If Not IsEmpty(c.Value) Then
5                 Title = "PeekCell (" + gAddinName + ")"
6                 TheValue = NonStringToString(c.Value2)
7                 Prompt = Prompt + AddressND(ActiveCell) & " value:" & vbLf & vbLf & TheValue
8                 If c.HasFormula Then
9                     TheFormula = GetFormula(c)
10                    Prompt = Prompt + vbLf + vbLf + AddressND(ActiveCell) + " formula:" + vbLf + vbLf + TheFormula
11                Else
12                    On Error Resume Next
13                    Set D = c.SpillParent
14                    On Error GoTo ErrHandler
15                    If Not D Is Nothing Then
16                        TheFormula = GetFormula(D)
17                        Prompt = Prompt + vbLf + vbLf + AddressND(ActiveCell) + " formula (spilling from " + AddressND(D) & "):" + vbLf + vbLf + TheFormula
18                    End If
19                End If
20                If TheFormula <> "" Then
21                    Res = MsgBoxPlus(Prompt, vbInformation + vbYesNoCancel + vbDefaultButton3, Title, "Copy &Value", "Copy &Formula", "OK", , 1200)
22                    If Res = vbYes Then
23                        CopyStringToClipboard TheValue
24                    ElseIf Res = vbNo Then
25                        CopyStringToClipboard TheFormula
26                    End If
27                Else
28                    Res = MsgBoxPlus(Prompt, vbInformation + vbOKCancel + vbDefaultButton2, Title, "Copy &Value", "OK", , , 1200)
29                    If Res = vbOK Then
30                        CopyStringToClipboard TheValue
31                    End If
32                End If
33            End If
34        End If

35        Exit Sub
ErrHandler:
36        SomethingWentWrong "#PeekCell (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function GetFormula(c As Range)
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
          Dim SPH As clsSheetProtectionHandler
          
3         If c.Parent.ProtectContents = False Or c.FormulaHidden = False Then
4             GetFormula = c.Formula2
5         ElseIf SheetIsProtectedWithPassword(c.Parent) Then
6             GetFormula = "Formula of cell " + AddressND(c) + " is hidden in a password-protected worksheet"
7         Else
8             Set SPH = CreateSheetProtectionHandler(c.Parent)
9             GetFormula = c.Formula2
10        End If

11        Exit Function
ErrHandler:
12        Throw "#GetFormula (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FlipNumberFormats
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Flips the NumberFormat on the current selection. Assigned to Ctrl+Shift+D.
'             We amend the order in which formats are tried according to the contents of
'             the active cell.
'             10-Nov-2017
'             If the active cell is a value displayed in a pivot table then the number
'             Format of the pivot field is changed
'             Assigned to Ctrl+Shift+D
' -----------------------------------------------------------------------------------------------------------------------
Sub FlipNumberFormats()
          Const fmt_date = "dd-mmm-yyyy"
          Const fmt_datetime = "dd-mmm-yyyy hh:mm:ss"
          Const fmt_gen = "General"
          Const fmt_money = "#,##0;[Red]-#,##0"
          Const fmt_moneypence = "#,##0.00;[Red]-#,##0.00"
          Dim FirstValue As Variant
          Dim FormatNew As String
          Dim FormatOld As String
          Dim Formats As Variant
          Dim LooksLikeADate As Boolean
          Dim LooksLikeADateTime As Boolean
          Dim MatchRes As Variant
          Dim o As Object
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler

2         EnsureAppObjectExists

3         If TypeName(Selection) <> "Range" Then Exit Sub
4         If Not UnprotectAsk(ActiveSheet, "Flip number formats") Then Exit Sub

5         On Error Resume Next
6         FirstValue = ActiveCell.Value2
7         If FirstValue >= CLng(CDate("1-Jan-1980")) Then
8             If FirstValue < CLng(CDate("1-Jan-2200")) Then
9                 If FirstValue = CLng(FirstValue) Then
10                    LooksLikeADate = True
11                Else
12                    LooksLikeADateTime = True
13                End If
14            End If
15        End If

16        On Error GoTo ErrHandler
17        If LooksLikeADateTime Then
18            Formats = sArrayStack(fmt_datetime, fmt_date, fmt_money, fmt_moneypence, fmt_gen)
19        ElseIf LooksLikeADate Then
20            Formats = sArrayStack(fmt_date, fmt_datetime, fmt_money, fmt_moneypence, fmt_gen)
21        Else
22            Formats = sArrayStack(fmt_money, fmt_moneypence, fmt_gen)
23        End If
24        FormatOld = Selection.Cells(1, 1).NumberFormat

25        MatchRes = sMatch(FormatOld, Formats)
26        If IsNumber(MatchRes) Then
27            If MatchRes = sNRows(Formats) Then
28                FormatNew = Formats(1, 1)
29            Else
30                FormatNew = Formats(MatchRes + 1, 1)
31            End If
32        Else
33            If LooksLikeADateTime Then
34                FormatNew = fmt_datetime
35            ElseIf LooksLikeADate Then
36                FormatNew = fmt_date
37            Else
38                FormatNew = "General"
39            End If
40        End If

41        For Each o In ObjectToSetNF()
42            If TypeName(o) = "PivotField" Then
                  'Because setting the NumberFormat of a PivotField triggers the sheet calculation event
43                If XSH Is Nothing Then Set XSH = CreateExcelStateHandler(, , False)
44                On Error Resume Next
45                o.NumberFormat = FormatNew
46                On Error GoTo ErrHandler
47            Else
48                o.NumberFormat = FormatNew
49            End If
50        Next o

51        Exit Sub
ErrHandler:
52        SomethingWentWrong "#FlipNumberFormats (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ObjectToSetNF
' Author    : Philip
' Date      : 10-Nov-2017
' Purpose   : Figure out the object that method FlipNumberFormats should operate on
'             actually returns a collection of objects...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ObjectToSetNF() As Object
          Dim clnRes As Collection
          Dim PC As PivotCell
          Dim PT As PivotTable
          Dim R As Range

          Dim pf As PivotField

1         On Error Resume Next
2         Set PC = ActiveCell.PivotCell
3         On Error GoTo ErrHandler
4         Set clnRes = New Collection
5         If PC Is Nothing Then
6             clnRes.Add Selection
7             Set ObjectToSetNF = clnRes
8             Exit Function
9         End If

10        Set PT = PC.Parent
11        For Each pf In PT.PivotFields
12            Set R = Nothing
13            On Error Resume Next
14            Set R = pf.DataRange
15            On Error GoTo ErrHandler
16            If Not R Is Nothing Then
17                If Not Application.Intersect(R, Selection) Is Nothing Then
18                    clnRes.Add pf
19                End If
20            End If
21        Next

22        For Each pf In PT.DataFields
23            Set R = Nothing
24            On Error Resume Next
25            Set R = pf.DataRange
26            On Error GoTo ErrHandler
27            If Not R Is Nothing Then
28                If Not Application.Intersect(R, Selection) Is Nothing Then
29                    clnRes.Add pf
30                End If
31            End If
32        Next

33        Set ObjectToSetNF = clnRes
34        Exit Function
ErrHandler:
35        Throw "#ObjectToSetNF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TogglePageBreaks
' Author    : Philip Swannell
' Date      : 02-Mar-2014
' Purpose   : Requested by Martin Baxter. Assigned to Ctrl+Shift+K
' -----------------------------------------------------------------------------------------------------------------------
Sub TogglePageBreaks()
1         On Error GoTo ErrHandler
          Dim b As Boolean
          Dim wb As Excel.Workbook
          Dim ws As Worksheet
2         If ActiveSheet Is Nothing Then Exit Sub
3         b = Not ActiveSheet.DisplayPageBreaks

4         For Each wb In Application.Workbooks
5             For Each ws In wb.Worksheets
6                 ws.DisplayPageBreaks = b
7             Next
8         Next
9         RefreshRibbon
10        Application.OnRepeat "Repeat Toggle Page Breaks", "TogglePageBreaks"

11        Exit Sub
ErrHandler:
12        MsgBoxPlus Err.Description, vbCritical
End Sub

Sub CtrlShiftI()
1         On Error GoTo ErrHandler
2         FormatAsInput
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#CtrlShiftI (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatAsInput
' Author    : Martin Baxter
' Date      : 14-Jun-2014
' Purpose   : Formats selection with blue text and unlocked.
'             Assigned to Ctrl+Shift+I
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatAsInput(Optional R As Range, Optional ByVal SwitchOn As String = "Toggle")
1         On Error GoTo ErrHandler

2         Select Case LCase$(SwitchOn)
              Case "on", "off", "toggle"
3             Case Else
4                 Throw "SwitchOn must be 'On', 'Off' or 'Toggle'"
5         End Select

6         If R Is Nothing Then
7             If TypeName(Selection) = "Range" Then
8                 Set R = Selection
9             Else
10                Exit Sub
11            End If
12        End If

13        If LCase$(SwitchOn) = "toggle" Then
              Dim a As Range
              Dim IsOn As Boolean
14            IsOn = True
15            For Each a In R.Areas
16                If IsNull(a.Font.Color) Or IsNull(a.Locked) Then
17                    IsOn = False
18                    Exit For
19                ElseIf a.Font.Color <> RGB(0, 0, 255) Or a.Locked <> False Then
20                    IsOn = False
21                    Exit For
22                End If
23            Next a
24            If IsOn Then
25                SwitchOn = "Off"
26            Else
27                SwitchOn = "On"
28            End If
29        End If

30        If Not UnprotectAsk(R.Parent, "Format as Input") Then Exit Sub
31        BackUpRange Selection, shUndo

32        If LCase$(SwitchOn) = "on" Then
33            R.Font.Color = RGB(0, 0, 255)
34            R.Locked = False
35        Else
36            R.Font.ColorIndex = xlColorIndexAutomatic
37            R.Locked = True
38        End If

39        If IsUndoAvailable(shUndo) Then
40            Application.OnUndo "Undo Format as Input", "RestoreRange"        'TODO Change RestoreRange to handle protected sheets
41        End If

42        Exit Sub
ErrHandler:
43        Throw "#FormatAsInput (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteValues
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Assigned to Ctrl+Shift+V
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteValues()
          Dim SPH As clsSheetProtectionHandler
          Dim UndoText As String
          Const Prompt = "Selection includes locked cells in a protected sheet. Are you sure you want to paste values?"
          Const Title = "Paste Values (Ctrl+Shift+V)"
          Dim ESH As clsExcelStateHandler
          Dim BiggerSelection As Range, Prompt2 As String, Prompt3 As String
          Dim Ask As Boolean

1         On Error GoTo ErrHandler

2         Ask = GetSetting(gAddinName, "PasteValues", "AskBeforePasting", "True")

3         If TypeName(Selection) = "Range" Then
4             Prompt3 = "Replace selected formulas with their values." + vbLf + vbLf + "Are you sure?" 'Tu request to have this dialog
5             If Ask Then
6                 Select Case MsgBoxPlus(Prompt3, vbYesNoCancel + vbQuestion + vbDefaultButton3, Title, "Yes, don't ask &Again", "Yes", "No, do nothing")
                      Case vbYes
7                         SaveSetting gAddinName, "PasteValues", "AskBeforePasting", "False"
8                     Case vbCancel
9                         Exit Sub
10                End Select
11            End If

12            Set BiggerSelection = ExpandRangeToIncludeEntireArrayFormulas(Selection)
13            If BiggerSelection.Cells.CountLarge <> Selection.CountLarge Then
14                Prompt2 = "Some array formulas were only partly included in the current selection, so the selection has been expanded." + vbLf + vbLf + _
                      "Paste values for the expanded selection?" + vbLf + vbLf + _
                      "Ctrl Z to undo."
15                BiggerSelection.Select
16                If MsgBoxPlus(Prompt2, vbOKCancel + vbQuestion, Title, "Paste values") <> vbOK Then
17                    Exit Sub
18                End If
19            End If

20            If sEquals(False, Selection.HasFormula) Then
21                If sEquals(False, Selection.HasArray) Then
22                    Throw "Cannot paste values because there are no formulas in the selection.", True        'Throw a benign error since user presumably imagined that there were formulas present.
23                End If
24            ElseIf Selection.Areas.Count > 1 Then
25                Throw "Cannot paste values on multiple selections.", True
26            End If

27            If Selection.Parent.ProtectContents Then
28                If Not sEquals(False, Selection.Locked) Then
29                    If Not SheetIsProtectedWithPassword(Selection.Parent) Then
30                        If MsgBoxPlus(Prompt, vbOKCancel + vbExclamation + vbDefaultButton2, Title, "Yes, Paste Values!") <> vbOK Then Exit Sub
31                        Set SPH = CreateSheetProtectionHandler(Selection.Parent)
32                    End If
33                End If
34            End If

35            Set ESH = CreateExcelStateHandler(, , , , , True)

36            BackUpRange Selection, shUndo

37            Selection.Copy
38            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

39            Application.CutCopyMode = False
40            UndoText = "Undo Paste Values at '" + AddressND(Selection) + "'"
41            If IsUndoAvailable(shUndo) Then
42                Application.OnUndo UndoText, "RestoreRange"        'TODO Change RestoreRange to handle protected sheets
43            End If
44            Application.OnRepeat "Repeat Paste Values", "PasteValues"
45        End If

46        Exit Sub
ErrHandler:
47        SomethingWentWrong "#PasteValues (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, Title
48        Application.CutCopyMode = False
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InsertFileNames
' Author    : Philip Swannell
' Date      : 25-Jun-2013
' Purpose   : Show the user a file selection dialog and paste the names of the files the
'             user selects into the active cell (and cells below when multiple files selected).
'             Attached to Ctrl+Shift+F and also available from the Ribbon.
' -----------------------------------------------------------------------------------------------------------------------
Sub InsertFileNames()
          Dim c As Range
          Dim DataToPaste As Variant
          Dim EscapeCharacter As String
          Dim Headers As Variant
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Overwrite As Boolean
          Dim RangeToPasteTo As Range
          Dim SUH As clsScreenUpdateHandler
          Dim Title As String

1         On Error GoTo ErrHandler

2         If ActiveCell Is Nothing Then Exit Sub
3         Title = "Insert File Names: Select File(s).    Hold down Shift Key to paste extra columns with file size, date etc."

4         DataToPaste = GetOpenFilenameWrap("InsertFileNames", , , _
              Title, "Paste", True, True, ActiveCell)
5         If VarType(DataToPaste) = vbBoolean Then Exit Sub
6         Set SUH = CreateScreenUpdateHandler()

7         Force2DArray DataToPaste

8         If sNCols(DataToPaste) > 1 Then DataToPaste = sArrayTranspose(DataToPaste)

9         If IsShiftKeyDown() Then
10            DataToPaste = AddMoreColumns(DataToPaste, Headers)
11            If VarType(DataToPaste) = vbBoolean Then Exit Sub
12        Else
13            Headers = sReshape("FullName", 1, 1)
14        End If

15        NR = sNRows(DataToPaste)
16        NC = sNCols(DataToPaste)

17        If NR + ActiveCell.row - 1 > ActiveCell.Parent.Rows.Count Then
18            Throw "Error: Target Range extends beyond the end of the worksheet", True
19        ElseIf NC + ActiveCell.Column - 1 > ActiveCell.Parent.Columns.Count Then
20            Throw "Error: Target Range extends beyond the end of the worksheet", True
21        End If

          'Prepend an escape character to stop Excel treating a string as not a _
           string, this step avoids problems with file names starting with \\, choose _
           the escape character to not change the horizontal alignment of the cell being pasted to
22        If False Then
23            For i = 1 To NR
24                For j = 1 To NC
25                    If VarType(DataToPaste(i, j)) = vbString Then
26                        Set c = ActiveCell.Cells(i, j)
27                        Select Case c.HorizontalAlignment
                              Case xlHAlignCenter
28                                EscapeCharacter = "^"
29                            Case xlHAlignRight
30                                EscapeCharacter = """"
31                            Case Else
32                                EscapeCharacter = "'"
33                        End Select
34                        DataToPaste(i, j) = EscapeCharacter & DataToPaste(i, j)
35                    End If
36                Next j
37            Next i
38        End If

39        Set RangeToPasteTo = ActiveCell.Resize(NR, NC)

40        If RangeToPasteTo.Cells.CountLarge = 1 Then
41            Overwrite = True
42        Else
43            Overwrite = True
44            For Each c In RangeToPasteTo.Cells
45                If Not IsEmpty(c.Value) Then
46                    Application.GoTo RangeToPasteTo
47                    Overwrite = MsgBoxPlus("Overwrite these cells?" + vbLf + vbLf + "(Ctrl Z to undo)", vbYesNo + vbQuestion + vbDefaultButton2, "Insert File Names") = vbYes
48                    Exit For
49                End If
50            Next c
51        End If

52        If Not UnprotectAsk(RangeToPasteTo.Parent, , RangeToPasteTo) Then Exit Sub

53        If Overwrite Then
54            BackUpRange RangeToPasteTo, shUndo
55            MyPaste RangeToPasteTo, DataToPaste
              'Hack for ISDA SIMM work
56            If Left(RangeToPasteTo.Parent.Parent.Name, 9) = "ISDA SIMM" Then
57                ISDASIMMInsertDataFolderCalls RangeToPasteTo, 0
58            End If

59            For i = 1 To NC
60                If Headers(1, i) = "Size" Or Headers(1, i) = "NumLines" Then
61                    RangeToPasteTo.Columns(i).NumberFormat = "###,##0"
62                ElseIf InStr(Headers(1, i), "Date") > 0 Then
63                    RangeToPasteTo.Columns(i).NumberFormat = "dd-mmm-yyyy hh:mm"
64                End If
65            Next i
66            Application.OnUndo "Undo Paste Filename" + IIf(NR > 1, "s", vbNullString) & " to " & AddressND(RangeToPasteTo), "RestoreRange"
67            RangeToPasteTo.Select
68        End If
69        Application.OnRepeat "Repeat Insert File Name", "InsertFileNames"
70        Exit Sub
ErrHandler:
71        SomethingWentWrong "#InsertFileNames (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, "Paste File Names"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddMoreColumns
' Author     : Philip Swannell
' Date       : 02-May-2018
' Purpose    : Sub of InsertFileNames - allow the user to select columns to be pasted to sheet
' Parameters :
'  FileNames: column array of file names with path
' -----------------------------------------------------------------------------------------------------------------------
Private Function AddMoreColumns(FileNames As Variant, ByRef Headers)
          Dim ButtonClicked As String
          Dim Info As Variant
          Dim InitialChoices
          Dim NotChosen
          Dim TheChoices As Variant
          Dim WithHeaders As Boolean
1         On Error GoTo ErrHandler

2         TheChoices = sTokeniseString("FullName,Name,DateCreated,DateLastAccessed,DateLastModified,Size,MD5,NumLines,Drive,Type,ParentFolder,ShortName,ShortPath,Attributes")
3         InitialChoices = sParseArrayString(GetSetting(gAddinName, "InsertFileNames", "Columns", "{""FullFileName""}"))
4         WithHeaders = GetSetting(gAddinName, "InsertFileNames", "WithHeaders", "False")
5         If Not sArraysIdentical(InitialChoices, sSubArray(TheChoices, 1, 1, sNRows(InitialChoices))) Then
6             NotChosen = sCompareTwoArrays(TheChoices, InitialChoices, "In1AndNotIn2")
7             If sNRows(NotChosen) > 1 Then
8                 TheChoices = sArrayStack(InitialChoices, sSubArray(NotChosen, 2))
9             Else
10                TheChoices = InitialChoices
11            End If
12        End If

13        Info = ShowMultipleChoiceDialog(TheChoices, InitialChoices, "Insert File Names", "Select columns to paste to sheet", , ActiveCell, , , False, , ButtonClicked, "With Headers", WithHeaders)
14        If ButtonClicked = "Cancel" Then
15            AddMoreColumns = False
16            Exit Function
17        End If
18        SaveSetting gAddinName, "InsertFileNames", "Columns", sMakeArrayString(Info)
19        SaveSetting gAddinName, "InsertFileNames", "WithHeaders", CStr(WithHeaders)
20        AddMoreColumns = sFileInfo(FileNames, sArrayTranspose(Info))
21        Headers = sArrayTranspose(Info)
22        If WithHeaders Then
23            AddMoreColumns = sArrayStack(Headers, AddMoreColumns)
24        End If

25        Exit Function
ErrHandler:
26        Throw "#AddMoreColumns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ModerniseWorkbook
' Author     : Philip Swannell
' Date       : 10-May-2018
' Purpose    : Takes a workbook that was originally created using a version of Excel that had Arial 10 point as the default font
'              and replaces all use of Arial with Calibri. Also makes the workbook's styles the same as the defaults for the current version of Excel.
'              Partly copied from https://support.microsoft.com/en-us/help/291321/how-to-programmatically-reset-a-workbook-to-default-styles
' Parameters :
'  MyBook: workbook to process, if omitted defaults to the active workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub ModerniseWorkbook(Optional MyBook As Excel.Workbook)

          Dim CurStyle As Style
          Dim i As Long
          Dim SPH() As clsSheetProtectionHandler
          Dim tempBook As Excel.Workbook
1         Application.ScreenUpdating = False

2         On Error GoTo ErrHandler
3         If MyBook Is Nothing Then Set MyBook = ActiveWorkbook

          'Unprotect all sheets in a way that will revert their protection status at method exit.
4         ReDim SPH(1 To MyBook.Worksheets.Count)
5         For i = 1 To MyBook.Worksheets.Count
6             Set SPH(i) = CreateSheetProtectionHandler(MyBook.Worksheets(i))
7         Next i

          'Delete all the styles in the workbook.
8         For Each CurStyle In MyBook.Styles
              'If CurStyle.Name <> "Normal" Then CurStyle.Delete
9             Select Case CurStyle.Name
                  Case "20% - Accent1", "20% - Accent2", _
                      "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
                      "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
                      "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
                      "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
                      "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
                      "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
                      "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
                      "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
                      "Note", "Output", "Percent", "Title", "Total", "Warning Text"
                      'Do nothing, these are the default styles
10                Case Else
11                    CurStyle.Delete
12            End Select

13        Next CurStyle

          'Open a new workbook.
14        Set tempBook = Workbooks.Add

          'Disable alerts so you may merge changes to the Normal style
          'from the new workbook.
15        Application.DisplayAlerts = False

          'Merge styles from the new workbook into the existing workbook.
16        MyBook.Styles.Merge Workbook:=tempBook

          'Enable alerts.
17        Application.DisplayAlerts = True

          'Close the new workbook.
18        tempBook.Close

          'Mmm - deleting old styles and merging in the default styles still leaves some cells formatted in Arial, so fix up
          Dim c As Range
          Dim ws As Worksheet
19        For Each ws In MyBook.Worksheets
20            For Each c In ws.UsedRange.Cells
21                If c.Font.Name = "Arial" Then
22                    c.Font.Name = "Calibri"
23                    If c.Font.Size = 10 Then c.Font.Size = 11
24                End If
25            Next c
26            IntersectWithComplement(ws.Range("$1:$1048576"), ws.UsedRange).Style = "Normal"
27        Next ws

28        Exit Sub
ErrHandler:
29        SomethingWentWrong "#ModerniseWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AutoFitColumns
' Author    : Philip Swannell
' Date      : 11-Apr-2016, moved to SolumAddin 24 April 2017
' Purpose   : Autofit the columns of a range of cells. ExtraWidth is an increase over
'             Excel's .AutoFit method. MinWidths passed as a row vector of minimum widths.
'             If MinWidths does not have enough elements then the last element is used for all remaining.
'             The function does not correctly handle ranges with more than one area, areas
'             other than the first area are processed by the .AutoFit command but subsequent
'             adjustments according to ExtraWidth, MinWidths, MaxWidth are not made.
' -----------------------------------------------------------------------------------------------------------------------
Sub AutoFitColumns(TheRange As Range, Optional ExtraWidth As Double, Optional ByVal MinWidths As Variant, Optional ByVal MaxWidths As Variant, Optional EmptyColWidth As Double)
          Dim HaveMax As Boolean
          Dim HaveMin As Boolean
          Dim i As Long
          Dim N As Long
          Dim NMax As Long
          Dim NMin As Long
          Dim ThisMax As Double
          Dim ThisMin As Double

1         On Error GoTo ErrHandler
2         HaveMax = Not IsMissing(MaxWidths)
3         HaveMin = Not IsMissing(MinWidths)
4         If HaveMin Then
5             Force2DArray MinWidths
6             NMin = sNCols(MinWidths)
7         End If
8         If HaveMax Then
9             Force2DArray MaxWidths
10            NMax = sNCols(MaxWidths)
11        End If
12        N = TheRange.Columns.Count
13        TheRange.Columns.AutoFit

14        If Not HaveMin Then If Not HaveMax Then If EmptyColWidth = 0 Then If ExtraWidth = 0 Then Exit Sub

15        For i = 1 To N
16            If IsColAllEmptyOrEmptyString(TheRange.Columns(i)) Then
17                If EmptyColWidth <> 0 Then
18                    TheRange.Columns(i).ColumnWidth = EmptyColWidth
19                End If
20            Else
21                If HaveMin Then
22                    ThisMin = MinWidths(1, IIf(i < NMin, i, NMin))
23                Else
24                    ThisMin = 0
25                End If
26                ThisMin = SafeMax(ThisMin, TheRange.Columns(i).ColumnWidth + ExtraWidth)
27                If ThisMin > 255 Then ThisMin = 255
28                If HaveMax Then
29                    ThisMax = MaxWidths(1, IIf(i < NMax, i, NMax))
30                End If
31                If TheRange.Columns(i).ColumnWidth < ThisMin Then TheRange.Columns(i).ColumnWidth = ThisMin
32                If HaveMax Then If TheRange.Columns(i).ColumnWidth > ThisMax Then TheRange.Columns(i).ColumnWidth = ThisMax
33            End If
34        Next i
35        Exit Sub
ErrHandler:
36        Throw "#AutoFitColumns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsColAllEmptyOrEmptyString
' Author    : Philip Swannell
' Date      : 21-Apr-2016
' Purpose   : AutoFit method does nothing on columns all of whose cells are empty or empty string
'             so we need a way of identifying such columns
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsColAllEmptyOrEmptyString(col As Range)
          Dim c As Range
          Dim i As Long
          Dim LastCell
          Dim SpecialCells As Range

1         On Error GoTo ErrHandler
2         If col.Cells.CountLarge = 1 Then
3             IsColAllEmptyOrEmptyString = (CStr(col.Value) = vbNullString)
4             Exit Function
5         End If

6         If Not IsEmpty(col.Cells(1, 1)) Then
7             If CStr(col.Cells(1, 1)) <> vbNullString Then
8                 IsColAllEmptyOrEmptyString = False
9                 Exit Function
10            End If
11        End If

12        If IsEmpty(col.Cells(1, 1)) Then
13            Set LastCell = col.Cells(1, 1).End(xlDown)
14            If LastCell.row >= col.row + col.Rows.Count Then
15                IsColAllEmptyOrEmptyString = True
16                Exit Function
17            ElseIf (LastCell.row = col.row + col.Rows.Count - 1) And (CStr(LastCell.Value) = vbNullString) Then
18                IsColAllEmptyOrEmptyString = True
19                Exit Function
20            End If
21        End If

22        For i = 1 To 2
23            Set SpecialCells = Nothing
24            On Error Resume Next
25            Set SpecialCells = col.SpecialCells(Choose(i, xlCellTypeConstants, xlCellTypeFormulas))
26            On Error GoTo 0

27            If Not SpecialCells Is Nothing Then
28                For Each c In SpecialCells.Cells
29                    If CStr(c.Value) <> vbNullString Then
30                        IsColAllEmptyOrEmptyString = False
31                        Exit Function
32                    End If
33                Next c
34            End If
35        Next i
36        IsColAllEmptyOrEmptyString = True
37        Exit Function
ErrHandler:
38        Throw "#IsColAllEmptyOrEmptyString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LocaliseGlobalNames
' Author    : Philip Swannell
' Date      : 07-Jun-2015
' Purpose   : Presents a dialog listing all names of the active workbook that refer to
'             ranges on the active sheet. User can select which ones should be scoped
'             only to the active sheet. Assigned to Ctrl + Alt + L.
' -----------------------------------------------------------------------------------------------------------------------
Sub LocaliseGlobalNames()
          Dim BookNames3Cols As Variant
          Dim ChooseVector As Variant
          Dim i As Long
          Dim j As Long
          Dim M As Name
          Dim N As Name
          Dim Res As Variant
          Dim SheetName As String
          Dim SheetNameWithQuotes As String
          Dim TargetNames As Variant
          Dim TheChoices As Variant
          Dim WorkbookNamesRefActiveSheet As Variant
          Const HeaderText = "Localise Global Names (" + gAddinName + ")"
          Dim TopText As String

1         On Error GoTo ErrHandler

2         If ActiveWorkbook Is Nothing Then
3             Throw "There is no active workbook.", True
4         ElseIf ActiveWorkbook.Names.Count = 0 Then
5             Throw "Active Workbook has no names.", True
6         End If

7         BookNames3Cols = sReshape(vbNullString, ActiveWorkbook.Names.Count, 3)
8         i = 0
9         For Each N In ActiveWorkbook.Names
10            i = i + 1
11            BookNames3Cols(i, 1) = N.Name
12            BookNames3Cols(i, 2) = N.RefersTo
13            BookNames3Cols(i, 3) = N.RefersToR1C1
14        Next N
15        SheetName = ActiveSheet.Name
16        SheetNameWithQuotes = "'" & Replace(SheetName, "'", "''") & "'"
17        ChooseVector = sArrayEquals("=" & SheetName & "!", sArrayLeft(sSubArray(BookNames3Cols, 1, 2, , 1), Len(SheetName) + 2))
18        ChooseVector = sArrayOr(ChooseVector, sArrayEquals("=" & SheetNameWithQuotes & "!", sArrayLeft(sSubArray(BookNames3Cols, 1, 2, , 1), Len(SheetNameWithQuotes) + 2)))
19        ChooseVector = sArrayAnd(ChooseVector, sArrayNot(sArrayFind("!", sSubArray(BookNames3Cols, 1, 1, , 1))))

20        If Not sColumnOr(ChooseVector)(1, 1) Then
21            Throw "No workbook-level names refer to ranges on the active sheet.", True
22        End If

23        WorkbookNamesRefActiveSheet = sMChoose(BookNames3Cols, ChooseVector)

24        TheChoices = sJustifyArrayOfStrings(sSubArray(WorkbookNamesRefActiveSheet, 1, 1, , 2), "Tahoma", 8, " " + vbTab)
25        TopText = "These workbook-level names refer to ranges on the active sheet. Select" + vbLf + _
              "names to ""localise"" i.e. names whose scope should be changed from the" + vbLf + _
              "entire workbook to just the active sheet." + vbLf + _
              vbNullString + vbLf + _
              "Warning: Formulas on other sheets that reference the names you localise" + vbLf + _
              "are likely to stop working and return the error #NAME?"

26        Res = ShowMultipleChoiceDialog(TheChoices, , HeaderText, TopText, False)
27        If IsEmpty(Res) Then GoTo EarlyExit
28        If sArraysIdentical(Res, "#User Cancel!") Then GoTo EarlyExit

29        ChooseVector = sArrayIsNumber(sMatch(TheChoices, Res))
30        TargetNames = sMChoose(WorkbookNamesRefActiveSheet, ChooseVector)
31        Force2DArray TargetNames

32        For i = 1 To sNRows(TargetNames)
              'Ensure that we get the book-level not sheet-level Name. Problem is when ActiveWorkbook.Names("Foo")
              'exists and Sheet1.Names("Foo") also exists. In this case ActiveWorkbook.Names("Foo") usually but not
              'always returns the book-level name. When it returns the sheet-level name we have to loop through all
              'names of the workbook to find the book-level name we are looking for.
33            Set M = ActiveWorkbook.Names(TargetNames(i, 1))
34            If M.Name = TargetNames(i, 1) And InStr(M.Name, "!") = 0 Then
                  'All good we found a book-level name
                  'Usually, deleting the book-level name will remove it. However, we may need to execute the deletion more than once!
35                For j = 1 To 10
36                    If IsInCollection(ActiveWorkbook.Names, CStr(TargetNames(i, 1))) Then
37                        Set M = ActiveWorkbook.Names(TargetNames(i, 1))
38                        If M.Name = TargetNames(i, 1) And InStr(M.Name, "!") = 0 Then
39                            M.Delete
40                        End If
41                    Else
42                        Exit For
43                    End If
44                Next j
45            Else
                  'Not good - we got a sheet-level name so we have to search all names in the workbook.
46                For Each M In ActiveWorkbook.Names
47                    If M.Name = TargetNames(i, 1) And InStr(M.Name, "!") = 0 Then
48                        M.Delete
49                        Exit For
50                    End If
51                Next M
52            End If
53            ActiveSheet.Names.Add Name:=TargetNames(i, 1), RefersToR1C1:=TargetNames(i, 3)
54        Next i
EarlyExit:
55        Application.OnRepeat "Repeat Localise Global Names", "LocaliseGlobalNames"

56        Exit Sub
ErrHandler:
57        SomethingWentWrong "#LocaliseGlobalNames (line " & CStr(Erl) + "): " & Err.Description & "!", , HeaderText
58        Application.OnRepeat "Repeat Localise Global Names", "LocaliseGlobalNames"
End Sub

Sub TestExecuteCommand()
          Dim WaitOnReturn As Boolean
1         On Error GoTo ErrHandler
2         WaitOnReturn = False
3         ExecuteCommand "C:\Program Files (x86)\Notepad++\notepad++.exe", "c:\temp\module1.bas", WaitOnReturn, vbMaximizedFocus
4         MsgBoxPlus Now
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#TestExecuteCommand (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExecuteCommand
' Author     : Philip Swannell
' Date       : 17-Aug-2020
' Purpose    : Run an executable with the option to wait until it's finished.
' Parameters :
'  Executable  : Full name of executable file. Maybe omitted in which case Arguments must carry the full syntax of the command to execute
'  Arguments   : Command line arguments to the executable
'  WaitOnReturn: If true then VBA execution waits until the executable file exits.
'  WindowStyle :
' -----------------------------------------------------------------------------------------------------------------------
Sub ExecuteCommand(Optional ByVal Executable As String, Optional Arguments As String, Optional WaitOnReturn As Boolean, Optional WindowStyle As VbAppWinStyle = vbMinimizedNoFocus)
          Dim wsh As WshShell
          Dim ErrorCode As Long

1         On Error GoTo ErrHandler
          Dim Command As String

2         If Executable <> "" Then
3             If Not sFileExists(Executable) Then Throw "Cannot find file '" + Executable + "'"
4             If InStr(Executable, " ") > 0 Then
5                 Command = Chr(34) + Executable + Chr(34)
6             Else
7                 Command = Executable
8             End If
9             If Arguments <> "" Then
10                Command = Command + " " + Arguments
11            End If
12        Else
13            Command = Arguments
14        End If

15        Set wsh = New WshShell

16        ErrorCode = wsh.Run(Command, WindowStyle, WaitOnReturn)

17        If ErrorCode <> 0 Then
18            Throw "Command '" + Command + "' failed with error code " + CStr(ErrorCode)
19        End If

20        Exit Sub
ErrHandler:
21        Throw "#ExecuteCommand (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
