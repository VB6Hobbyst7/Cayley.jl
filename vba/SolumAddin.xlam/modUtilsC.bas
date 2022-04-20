Attribute VB_Name = "modUtilsC"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReleasesPerMonth
' Author     : Philip Swannell
' Date       : 17-Oct-2019
' Purpose    : Used on the Audit sheet of this workbook to generate data for graph showing number of releases of SolumAddin per month.
' -----------------------------------------------------------------------------------------------------------------------
Function ReleasesPerMonth()
          Dim RawData, MaxDate, MinDate, NumMonths, ExtraDates, i As Long
1         On Error GoTo ErrHandler
2         RawData = sExpandDown(shAudit.Range("Headers").Cells(2, 2)).Value2
3         MaxDate = sMaxOfArray(RawData)
4         MinDate = sMinOfArray(RawData)
5         NumMonths = 12 * (Year(MaxDate) - Year(MinDate)) + Month(MaxDate) - Month(MinDate) + 1
6         ExtraDates = sReshape(0, NumMonths, 1)
7         For i = 1 To NumMonths
8             ExtraDates(i, 1) = CLng(DateSerial(Year(MinDate), Month(MinDate) - 1 + i, 1))
9         Next
10        RawData = sArrayStack(RawData, ExtraDates)
11        RawData = sSortedArray(RawData)
12        For i = 1 To sNRows(RawData)
13            RawData(i, 1) = Format$(RawData(i, 1), "mmm-yy")
14        Next
15        RawData = sCountRepeats(RawData, "CH")
16        For i = 1 To sNRows(RawData)
17            RawData(i, 2) = RawData(i, 2) - 1
18        Next
19        ReleasesPerMonth = RawData
20        Exit Function
ErrHandler:
21        ReleasesPerMonth = "#ReleasesPerMonth (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateSheetProtectionHandler
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Method to simplify the handling of worksheet protection.
'             As this method runs (and the class is created) the protection of the sheet
'             (DrawingObjects, Contents and Scenarios) is set to SetStateTo. When the object goes out of scope the
'             protection of the sheet is set to how it was at the time the class was instantiated.
' -----------------------------------------------------------------------------------------------------------------------
Function CreateSheetProtectionHandler(ws As Worksheet, _
        Optional SetStateTo As Boolean = False, _
        Optional Password As String) As clsSheetProtectionHandler
1         On Error GoTo ErrHandler

2         Set CreateSheetProtectionHandler = New clsSheetProtectionHandler
3         CreateSheetProtectionHandler.Init ws, SetStateTo, Password

4         Exit Function
ErrHandler:
5         Throw "#CreateSheetProtectionHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePositionInstructions
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Creates a clsPositionInstructions object that can be interpreted by both
'             SetFormPosition to position a form and by ShowCommandBarPopUp to position
'             a command bar.
' -----------------------------------------------------------------------------------------------------------------------
Sub CreatePositionInstructions(ByRef PI As clsPositionInstructions, AnchorObject As Object, X_Nudge As Double, Y_Nudge As Double)
1         On Error GoTo ErrHandler
2         Set PI = New clsPositionInstructions
3         Set PI.AnchorObject = AnchorObject
4         PI.X_Nudge = X_Nudge
5         PI.Y_Nudge = Y_Nudge
6         Exit Sub
ErrHandler:
7         Throw "#CreatePositionInstructions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub CreateFormResizer(FR As clsFormResizer)
1         Set FR = New clsFormResizer
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateStacker
' Author    : Philip Swannell
' Date      : 09-Jun-2015
' Purpose   : So that we can create clsStacker objects from other workbooks...
' -----------------------------------------------------------------------------------------------------------------------
Function CreateStacker() As clsStacker
1         On Error GoTo ErrHandler
2         Set CreateStacker = New clsStacker
3         Exit Function
ErrHandler:
4         Throw "#CreateStacker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateHStacker
' Author    : Philip Swannell
' Date      : 24-Aug-2020
' Purpose   : So that we can create clsHStacker objects from other workbooks...
' -----------------------------------------------------------------------------------------------------------------------
Function CreateHStacker() As clsHStacker
1         On Error GoTo ErrHandler
2         Set CreateHStacker = New clsHStacker
3         Exit Function
ErrHandler:
4         Throw "#CreateHStacker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub TestHStacker()
          Dim STK As clsHStacker
          Dim i As Long

1         On Error GoTo ErrHandler
2         Set STK = CreateHStacker()

3         For i = 1 To 10
4             STK.StackData Array(i, i + 1, i + 2)
5         Next
6         STK.StackData sReshape(sIntegers(9), 3, 10)
7         g STK.Report
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#TestHStacker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateScreenUpdateHandler
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Makes it easy to control screen updating when there are long call stacks.
' -----------------------------------------------------------------------------------------------------------------------
Function CreateScreenUpdateHandler(Optional SetStateTo As Boolean = False) As clsScreenUpdateHandler
1         On Error GoTo ErrHandler
2         Set CreateScreenUpdateHandler = New clsScreenUpdateHandler
3         CreateScreenUpdateHandler.Init SetStateTo
4         Exit Function
ErrHandler:
5         Throw "#CreateScreenUpdateHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateExcelStateHandler
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : See comments at the top of clsExcelStateHandler
' -----------------------------------------------------------------------------------------------------------------------
Function CreateExcelStateHandler(Optional SetCalculationTo As XlCalculation, _
        Optional SetReferenceStyleTo As XlReferenceStyle, _
        Optional SetEnableEventsTo As Variant, _
        Optional SetStatusBarTo As String, _
        Optional SetEditDirectlyInCellTo As Variant, _
        Optional PreserveViewport As Boolean) As clsExcelStateHandler
1         On Error GoTo ErrHandler

2         Set CreateExcelStateHandler = New clsExcelStateHandler
3         CreateExcelStateHandler.Init SetCalculationTo, SetReferenceStyleTo, _
              SetEnableEventsTo, SetStatusBarTo, SetEditDirectlyInCellTo, PreserveViewport

4         Exit Function
ErrHandler:
5         Throw "#CreateExcelStateHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsInCollection
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Tests for membership of any collection. Can be used in place of multiple
'             functions BookHasSheet, SheetHasName etc. Works irrespective of whether
'             the collection contains objects or primitives.
' -----------------------------------------------------------------------------------------------------------------------
Public Function IsInCollection(oColn As Object, Key As String) As Boolean
1         On Error GoTo ErrHandler
2         VarType (oColn(Key))
3         IsInCollection = True
4         Exit Function
ErrHandler:
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AssignKeys
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Add all assigned keys here for switch on when add-in opens and switch off
'             before it closes. Unfortunately MZTools key assignments conflict with these
'             even though MZTools assignments should only apply in the VBE.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub AssignKeys(SwitchOn As Boolean)
          'Shift = +
          'Ctrl = ^
          'Alt = %
          Dim M As Variant
          Dim wbn As String
1         On Error GoTo ErrHandler

2         If ActiveWindow Is Nothing Then
3             If Application.ProtectedViewWindows.Count > 0 Then
4                 Exit Sub
5             End If
6         End If

7         M = CreateMissing()
8         wbn = ThisWorkbook.Name
9         If ExcelSupportsSpill() Then
10            Application.OnKey "^/", IIf(SwitchOn, wbn & "!CtrlForwardslashResponse", M)
              'Experimental feature 4 Dec 2019. Stop user from entering old-style CSE array formulas by immediately converting them to new-style dynamic arrays!
              'switched off 5 Feb 2020. Doesn't play well with Bloomberg function BDH.
              ' Application.OnKey "^+{ENTER}", IIf(SwitchOn, wbn & "!FitArrayFormula", m)
              'Application.OnKey "^+~", IIf(SwitchOn, wbn & "!FitArrayFormula", m)
11        End If
12        Application.OnKey "%`", IIf(SwitchOn, wbn & "!AltBacktickResponse", M)
13        Application.OnKey "^%+c", IIf(SwitchOn, wbn & "!CalcActiveSheet", M)
14        Application.OnKey "^%+f", IIf(SwitchOn, wbn & "!InsertFolderName", M)          'documented
15        Application.OnKey "^%+g", IIf(SwitchOn, wbn & "!AddGreyBordersAround", M)
16        Application.OnKey "^%+w", IIf(SwitchOn, wbn & "!ArrangeWindows", M)
17        Application.OnKey "^%b", IIf(SwitchOn, wbn & "!KeyboardAccessToButtons", M)    'documented
18        Application.OnKey "^%f", IIf(SwitchOn, wbn & "!SearchWorkbookFormulas", M)     'documented
19        Application.OnKey "^%g", IIf(SwitchOn, wbn & "!AddGreyBorders", M)
20        Application.OnKey "^%l", IIf(SwitchOn, wbn & "!LocaliseGlobalNames", M)        'documented
21        Application.OnKey "^%m", IIf(SwitchOn, wbn & "!ShowMessages", M)               'documented
22        Application.OnKey "^{F6}", IIf(SwitchOn, wbn & "!CtrlF6Response", M)
23        Application.OnKey "^+{INSERT}", IIf(SwitchOn, wbn & "!PasteDuplicateRange", M) 'documented
24        Application.OnKey "^+a", IIf(SwitchOn, wbn & "!ResizeArrayFormula", M)         'documented
25        Application.OnKey "^+b", IIf(SwitchOn, wbn & "!SwitchBook", M)                 'documented
26        Application.OnKey "^+c", IIf(SwitchOn, wbn & "!CalcSelection", M)              'documented
27        Application.OnKey "^+d", IIf(SwitchOn, wbn & "!FlipNumberFormats", M)          'documented
28        Application.OnKey "^+f", IIf(SwitchOn, wbn & "!InsertFileNames", M)            'documented
29        Application.OnKey "^+g", IIf(SwitchOn, wbn & "!ExtractArguments", M)           'documented
30        Application.OnKey "^+i", IIf(SwitchOn, wbn & "!CtrlShiftI", M)                 'documented
31        Application.OnKey "^+j", IIf(SwitchOn, wbn & "!JustifyText", M)                'documented
32        Application.OnKey "^+k", IIf(SwitchOn, wbn & "!TogglePageBreaks", M)           'documented
33        Application.OnKey "^+m", IIf(SwitchOn, wbn & "!WhiteBoldOnBlue", M)
34        Application.OnKey "^+r", IIf(SwitchOn, wbn & "!FitArrayFormula", M)            'documented
35        Application.OnKey "^+t", IIf(SwitchOn, wbn & "!SwitchSheet", M)                'documented
36        Application.OnKey "^+v", IIf(SwitchOn, wbn & "!PasteValues", M)                'documented
37        Application.OnKey "^+w", IIf(SwitchOn, wbn & "!ToggleWindow", M)               'documented
38        Application.OnKey "^2", IIf(SwitchOn, wbn & "!PastePictureToActiveCell", M)
39        Application.OnKey "^n", IIf(SwitchOn, wbn & "!AddNewWorkBook", M)
40        Application.OnKey "^q", IIf(SwitchOn, wbn & "!QuitExcel", M)
41        Application.OnKey "{F12}", IIf(SwitchOn, wbn & "!PeekCell", M)

          'Line below very handy when using Camtasia
          'Application.OnKey "^{CAPSLOCK}", IIf(SwitchOn, wbn & "!ResizeWindowsForVideoRecording", M)
42        SwitchF10ResponseOn

43        Exit Sub
ErrHandler:
44        Throw "#AssignKeys (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sJustifyText
' Author    : Philip Swannell
' Date      : 11-Oct-2016
' Purpose   : Justifies text like Excel's Home > Fill > Justify, but as a spreadsheet function.
' Arguments
' CurrentText: A column of text. Non-text elements will be cast as text except for empty elements (blank
'             cells) which are taken to indicate the start of a new paragraph.
' FontName  : The name of the font such as "Calibri".
' FontSize  : The size of the font in points. E.g. Excel defaults to a size 11 Calibri font.
' AvailableWidth: When displayed in the given font, the returned column of text will have width in points no
'             greater than AvailableWidth, except that individual words of greater width
'             appear in the return without being broken or hyphenated.
' -----------------------------------------------------------------------------------------------------------------------
Function sJustifyText(ByVal CurrentText, FontName As String, FontSize As Long, AvailableWidth As Double)
Attribute sJustifyText.VB_Description = "Justifies text like Excel's Home > Fill > Justify, but as a spreadsheet function."
Attribute sJustifyText.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim BlankIndicators
          Dim CountRepeatsRes
          Dim i As Long
          Dim STK As clsStacker

1         On Error GoTo ErrHandler
2         Force2DArrayR CurrentText
3         BlankIndicators = sReshape(True, sNRows(CurrentText), 1)
4         For i = 1 To sNRows(CurrentText)
5             BlankIndicators(i, 1) = IsEmpty(CurrentText(i, 1))
6         Next i
7         CountRepeatsRes = sCountRepeats(BlankIndicators, "CFH")

8         Set STK = CreateStacker()
9         For i = 1 To sNRows(CountRepeatsRes)
10            If CountRepeatsRes(i, 1) Then
11                STK.Stack2D sReshape(Empty, CLng(CountRepeatsRes(i, 3)), 1)
12            Else
13                STK.StackData CoreJustifyText(sSubArray(CurrentText, CountRepeatsRes(i, 2), 1, CountRepeatsRes(i, 3)), FontName, FontSize, AvailableWidth)
14            End If
15        Next i

16        sJustifyText = STK.Report

17        Exit Function
ErrHandler:
18        sJustifyText = "#sJustifyText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreJustifyText
' Author    : Philip Swannell
' Date      : 11-Oct-2016
' Purpose   : Sub-routine of sJustifyText. Handles "one paragraph".
' -----------------------------------------------------------------------------------------------------------------------
Private Function CoreJustifyText(ByVal CurrentText, FontName As String, FontSize As Long, AvailableWidth As Double)
          Dim CheckString1 As String
          Dim CheckString2 As String
          Dim i As Long
          Dim spaceWidth As Double
          Dim STK As clsStacker
          Dim ThisLine As String
          Dim ThisLineLength As Double
          Dim ThisWordCount As Long
          Dim WordLengths
          Dim Words

1         On Error GoTo ErrHandler

2         Force2DArray CurrentText

3         CurrentText = sConcatenateStrings(CurrentText, " ")
4         CheckString1 = CurrentText
5         Words = sTokeniseString(CStr(CurrentText), " ")
6         WordLengths = sStringWidth(Words, FontName, FontSize)
7         Force2DArrayRMulti Words, WordLengths
8         Set STK = CreateStacker()

9         spaceWidth = sStringWidth(" ", FontName, FontSize)(1, 1)
10        For i = 1 To sNRows(Words)
11            If ThisWordCount = 0 Or ThisLineLength + spaceWidth + WordLengths(i, 1) < AvailableWidth Then
12                ThisLine = ThisLine + IIf(ThisWordCount = 0, vbNullString, " ") + Words(i, 1)
13                ThisLineLength = ThisLineLength + IIf(ThisWordCount = 0, 0, spaceWidth) + WordLengths(i, 1)
14                ThisWordCount = ThisWordCount + 1
15            Else
16                STK.StackData ThisLine
17                ThisLine = Words(i, 1)
18                ThisLineLength = WordLengths(i, 1)
19                ThisWordCount = 1
20            End If
21        Next i
22        STK.StackData ThisLine

23        CoreJustifyText = STK.Report
24        CheckString2 = sConcatenateStrings(CoreJustifyText, " ")

25        If CheckString1 <> CheckString2 Then
26            Throw "Assertion Failed. Proposed justified text differs from the starting text."
27        End If

28        Exit Function
ErrHandler:
29        Throw "#CoreJustifyText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RangeFromSheet
' Author    : Philip Swannell
' Date      : 10-Jun-2016
' Purpose   : Use instead of sh.Range(RangeName) so that errors are more informative than
'             "Method 'Range' of object '_Worksheet' failed". Additionally the method can
'             validate the contents of the range according to the Allow... arguments.
' -----------------------------------------------------------------------------------------------------------------------
Function RangeFromSheet(sh As Worksheet, RangeName As String, Optional AllowNumbers As Boolean = True, _
        Optional AllowStrings As Boolean = True, Optional AllowBooleans As Boolean = True, _
        Optional AllowEmpty As Boolean = True, Optional AllowErrors As Boolean = True) As Range
          
          Dim c As Range
          Dim ErrorMessage As String
          Dim ErrorMessage2 As String
          Dim g1 As Boolean
          Dim i As Long
          Dim j As Long
          Dim NumAllowedTypes As Long
          Dim R As Range
          
1         On Error Resume Next
2         Set R = sh.Range(RangeName)
3         On Error GoTo ErrHandler

4         If R Is Nothing Then Throw "Cannot find range named '" + RangeName + "' on worksheet '" + sh.Name + "' of workbook '" + sh.Parent.Name + "'"
5         If Not R.Parent Is sh Then Throw "Detected incorrect range definition: worksheet '" + sh.Name + "' of workbook '" + sh.Parent.Name + "' has a name '" + RangeName + "' but thnat name refers to a range on a different worksheet."
6         Set RangeFromSheet = R
7         If (AllowNumbers And AllowStrings And AllowBooleans And AllowEmpty And AllowErrors) Then Exit Function
8         If Not (AllowNumbers Or AllowStrings Or AllowBooleans Or AllowEmpty Or AllowErrors) Then Throw "At least one of AllowNumbers, AllowStrings, AllowBooleans,AllowEmpty or AllowErrors must be TRUE"
9         g1 = R.Cells.CountLarge > 1

10        For Each c In R.Cells
11            Select Case VarType(c.Value2)
                  Case vbString
12                    If Not AllowStrings Then
13                        ErrorMessage2 = IIf(g1, " but cell " + AddressND(c), "  but it") + " contains text"
14                        GoTo Fail
15                    End If
16                Case vbError
17                    If Not AllowErrors Then
18                        ErrorMessage2 = IIf(g1, " but cell " + AddressND(c), " but it") + " contains an error"
19                        GoTo Fail
20                    End If
21                Case vbBoolean
22                    If Not AllowBooleans Then
23                        ErrorMessage2 = IIf(g1, " but cell " + AddressND(c), " but it") + " contains a logical value"
24                        GoTo Fail
25                    End If
26                Case vbEmpty
27                    If Not AllowEmpty Then
28                        ErrorMessage2 = IIf(g1, " but cell " + AddressND(c), " but it") + " is empty"
29                        GoTo Fail
30                    End If
31                Case Else
32                    If Not AllowNumbers Then
33                        ErrorMessage2 = IIf(g1, " but cell " + AddressND(c), " but it") + " contains a number"
34                        GoTo Fail
35                    End If
36            End Select
37        Next c
38        Exit Function
Fail:

39        If UCase$(Replace(RangeName, "$", vbNullString)) <> UCase$(AddressND(R)) Then
40            ErrorMessage = "'" + RangeName + "' (" + AddressND(R) + ")"
41        Else
42            ErrorMessage = RangeName
43        End If

44        ErrorMessage = ErrorMessage + " on worksheet '" + sh.Name + "' of workbook '" + sh.Parent.Name + "' must be"

45        If g1 Then
46            ErrorMessage = "Every cell in range " + ErrorMessage
47        Else
48            ErrorMessage = "Range " + ErrorMessage
49        End If
50        NumAllowedTypes = -(AllowStrings + AllowNumbers + AllowBooleans + AllowEmpty + AllowErrors)
51        If NumAllowedTypes > 1 Then
52            ErrorMessage = ErrorMessage + " either"
53        End If
54        For i = 1 To 5
55            If Choose(i, AllowNumbers, AllowStrings, AllowBooleans, AllowEmpty, AllowErrors) Then
56                j = j + 1
57                ErrorMessage = ErrorMessage + IIf(NumAllowedTypes > 1 And j > 1, " or ", " ") + Choose(i, "a number", "text", "TRUE or FALSE", "empty", "an error")
58            End If
59        Next i
60        Throw ErrorMessage + ErrorMessage2

61        Exit Function
ErrHandler:
62        Throw "#RangeFromSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CompareWorkbookFormulas
' Author    : Philip Swannell
' Date      : 28-Jan-2016
' Purpose   : Compare formulas in two workbooks, code currently ad-hoc and needs to be
'             edited each time it's used. Should work up into something more user friendly
' -----------------------------------------------------------------------------------------------------------------------
Sub CompareWorkbookFormulas()
          Const IgnoreTheseSheets = vbNullString        '",Portfolio,xvaDashboard,Regression,HiddenSheet,Audit,Config,CounterpartyViewer,TradeViewer,"
          Dim C1 As Range
          Dim C2 As Range
          Dim i As Long
          Dim RangeToProcess As Range
          Dim ReportSheet As Worksheet
          Dim SPH1 As clsSheetProtectionHandler
          Dim SPH2 As clsSheetProtectionHandler
          Dim UsedRange1 As Range
          Dim UsedRange2 As Range
          Dim wb1 As Excel.Workbook
          Dim wb2 As Excel.Workbook
          Dim ws1 As Worksheet
          Dim ws2 As Worksheet

1         On Error GoTo ErrHandler
2         Application.ScreenUpdating = False

3         i = 2
4         Set wb1 = Application.Workbooks("SCRiPT10Jan.xlsm")
5         Set wb2 = Application.Workbooks("SCRiPT.xlsm")
6         Set ReportSheet = Application.Workbooks.Add.Worksheets(1)
7         ReportSheet.Cells(i, 1).Value = "Sheet Name"
8         ReportSheet.Cells(i, 2).Value = "Cell Address"
9         ReportSheet.Cells(i, 3).Value = "Book1 Array Address"
10        ReportSheet.Cells(i, 4).Value = "Book2 Array Address"

11        ReportSheet.Cells(i, 5).Value = "Book1 Formula"
12        ReportSheet.Cells(i, 6).Value = "Book2 Formula"

          Dim DoThisOne As Boolean

13        For Each ws1 In wb1.Worksheets
14            Set SPH1 = CreateSheetProtectionHandler(ws1)
15            If InStr(LCase$(IgnoreTheseSheets), LCase$("," + ws1.Name + ",")) = 0 Then
16                If IsInCollection(wb2.Worksheets, ws1.Name) Then
17                    Set ws2 = wb2.Worksheets(ws1.Name)
18                    Set SPH2 = CreateSheetProtectionHandler(ws2)
19                    Set UsedRange1 = ws1.UsedRange
20                    Set UsedRange2 = ws2.UsedRange
21                    Set RangeToProcess = Application.Union(UsedRange1, ws1.Range(UsedRange2.address))
22                    For Each C1 In RangeToProcess.Cells
23                        Set C2 = ws2.Range(C1.address)
24                        If C1.Formula <> C2.Formula Then
25                            DoThisOne = True
                              '  If LCase(c1.Formula) = LCase(c2.Formula) Then DoThisOne = False
26                            If Left$(C1.Formula, 1) <> "=" And Left$(C2.Formula, 1) <> "=" Then DoThisOne = False        'don't want to compare two values
                              ' If c1.HasArray Then If c1.Address <> c1.CurrentArray.Cells(1, 1).Address Then DoThisOne = False
                              ' If c2.HasArray Then If c2.Address <> c2.CurrentArray.Cells(1, 1).Address Then DoThisOne = False
27                            If DoThisOne Then
28                                i = i + 1
29                                ReportSheet.Cells(i, 1).Value = ws1.Name
30                                ReportSheet.Cells(i, 2).Value = C1.address
31                                If C1.HasArray Then
32                                    ReportSheet.Cells(i, 3).Value = C1.CurrentArray.address
33                                End If
34                                If C2.HasArray Then
35                                    ReportSheet.Cells(i, 4).Value = C2.CurrentArray.address
36                                End If

37                                ReportSheet.Cells(i, 5).Value = "'" + C1.Formula
38                                ReportSheet.Cells(i, 6).Value = "'" + C2.Formula
39                            End If
40                        End If
41                    Next C1
42                End If
43            End If
44        Next ws1

45        Exit Sub
ErrHandler:
46        SomethingWentWrong "#CompareWorkbookFormulas (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SheetIsProtectedWithPassword
' Author    : Philip Swannell
' Date      : 03-Nov-2015
' Purpose   : Returns True if the Contents are protected and there is a password. Leaves
'             the protection state of the worksheet unchanged.
' -----------------------------------------------------------------------------------------------------------------------
Function SheetIsProtectedWithPassword(ws As Worksheet) As Boolean
1         On Error GoTo ErrHandler
2         If ws.ProtectContents = False Then
3             SheetIsProtectedWithPassword = False
4         Else
5             On Error Resume Next
6             ws.Unprotect "UnlikelyPassword" + CStr(Rnd())
7             On Error GoTo ErrHandler
8             If ws.ProtectContents Then
9                 SheetIsProtectedWithPassword = True
10            Else
11                ws.Protect , , True
12                SheetIsProtectedWithPassword = False
13            End If
14        End If
15        Exit Function
ErrHandler:
16        Throw "#SheetIsProtectedWithPassword (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddGreyBordersAround
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Attached to Ctrl Alt Shift G
' -----------------------------------------------------------------------------------------------------------------------
Sub AddGreyBordersAround()
1         AddGreyBorders , True
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddGreyBorders
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Inside and edge borders light-grey. The grey chosen is indicated in Excel's
'             colour-picker as "White, Background 1, Darker 25%"
'             Attached to Ctrl+Alt+G
' -----------------------------------------------------------------------------------------------------------------------
Sub AddGreyBorders(Optional TheRange As Range, Optional ByVal EdgesOnly As Boolean)
          Dim i As Long
          Const Title = "Make Borders Light Grey"

1         On Error GoTo ErrHandler
          Dim SUH As clsScreenUpdateHandler
2         Set SUH = CreateScreenUpdateHandler()

3         If TheRange Is Nothing Then
4             If TypeName(Selection) <> "Range" Then Throw "You must select a range to set the borders", True
5             Set TheRange = Selection
6         End If

7         If Not UnprotectAsk(TheRange.Parent, Title) Then Exit Sub

8         With TheRange
9             For i = 1 To 8
10                .Borders(Choose(i, xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)).LineStyle = xlNone
11            Next i
12            For i = 1 To IIf(EdgesOnly, 4, 6)
13                With .Borders(Choose(i, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal))
14                    .LineStyle = xlContinuous
15                    .Weight = xlThin
16                    .ThemeColor = 1
17                    .TintAndShade = -0.249946592608417
18                End With
19            Next i
20        End With
21        Application.OnRepeat "Repeat Grey Borders", IIf(EdgesOnly, "AddGreyBordersAround", "AddGreyBorders")
22        Exit Sub
ErrHandler:
23        SomethingWentWrong "#AddGreyBorders (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : WhiteBoldOnBlue
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Thu likes to Format "Header cells" in Solum Blue - this macro does it...
' -----------------------------------------------------------------------------------------------------------------------
Sub WhiteBoldOnBlue()
          Const Title = "Solum Blue"
          Dim c As Range
1         On Error GoTo ErrHandler
2         If TypeName(Selection) <> "Range" Then Throw "No cells are selected", True
3         If Not UnprotectAsk(ActiveSheet, Title) Then Exit Sub

4         BackUpRange Selection, shUndo

5         With Selection
6             .Font.Bold = True
7             .Font.Color = RGB(255, 255, 255)
8             If .Rows.Count = 1 Then    ' i.e. we are formatting a header row
9                 For Each c In .Offset(0).Cells
10                    If Not IsEmpty(c.Offset(1).Value) Then
11                        c.HorizontalAlignment = c.Offset(1).HorizontalAlignment
12                    End If
13                Next
14            End If
15            .VerticalAlignment = xlVAlignCenter
16            .Interior.Color = RGB(0, 102, 204)        'Solum's corporate colour
17        End With

18        If IsUndoAvailable(shUndo) Then
19            Application.OnUndo "Undo " & Title, "RestoreRange"      'TODO Change RestoreRange to handle protected sheets
20        End If
21        Application.OnRepeat "Repeat " & Title, "WhiteBoldOnBlue"

22        Exit Sub
ErrHandler:
23        SomethingWentWrong "#WhiteBoldOnBlue (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuitExcel
' Author    : Philip Swannell
' Date      : 11-Dec-2015
' Purpose   : Attached to Ctrl Q plus available from the backstage (fka File menu)
' -----------------------------------------------------------------------------------------------------------------------
Sub QuitExcel()
1         Application.Quit
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFontIsInstalled
' Author     : Philip Swannell
' Date       : 20-Nov-2018
' Purpose    : Test whether font is available. Unfortunately the method CoreFontIsInstalled only works in
'              certain circumstances - e.g. there is an active window, so we cache the values of the return for
'              for those fonts that we call this function for. See also method RecordInstalledFonts, called indirectly from Workbook_Open
' -----------------------------------------------------------------------------------------------------------------------
Function sFontIsInstalled(FontName As String)
Attribute sFontIsInstalled.VB_Description = "Returns TRUE if a font with FontName is installed on the PC, FALSE otherwise."
Attribute sFontIsInstalled.VB_ProcData.VB_Invoke_Func = " \n28"
1         Select Case LCase$(FontName)
              Case "wingdings", "wingdings 3"
2                 sFontIsInstalled = shAudit.Range("FontInstalled" & Replace(FontName, " ", "_")).Value
3             Case Else
4                 sFontIsInstalled = CoreFontIsInstalled(FontName)
5         End Select
End Function

Sub RecordInstalledFonts()
1         On Error GoTo ErrHandler
          Dim FontName As String
          Dim i As Long
          Dim wb As Excel.Workbook
2         If Application.ScreenUpdating Then Application.ScreenUpdating = False
3         If ActiveWindow Is Nothing Then Set wb = Application.Workbooks.Add
4         For i = 1 To 2
5             FontName = Choose(i, "Wingdings", "Wingdings 3")
6             shAudit.Range("FontInstalled" & Replace(FontName, " ", "_")) = CoreFontIsInstalled(FontName)
7         Next
8         If Not wb Is Nothing Then wb.Close False
9         ThisWorkbook.Saved = True

10        Exit Sub
ErrHandler:
11        SomethingWentWrong "#RecordInstalledFonts (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JustifyText
' Author    : Philip Swannell
' Date      : 09-Oct-2016
' Purpose   : Replacement for Excel's Home > Fill > Justify but without the annoying 255
'             character limit. Assigned to Ctrl+Shift+J
' -----------------------------------------------------------------------------------------------------------------------
Sub JustifyText()
          Dim AvailableWidth As Double
          Dim CopyFrom As Range
          Dim BottomCell As Range
          Dim c As Range
          Dim CurrentText As Variant
          Dim FontName As Variant
          Dim FontSize As Variant
          Dim JustifiedText As Variant
          Dim OldSelection As Range
          Dim Prompt As String
          Dim SUH As clsScreenUpdateHandler
          Dim TargetRange As Range
          Const Title = "Justify Text"

1         On Error GoTo ErrHandler
2         If TypeName(Selection) <> "Range" Then Exit Sub
3         If Selection.Areas.Count > 1 Then Throw "This action won't work on multiple selections.", True
4         If Not UnprotectAsk(ActiveSheet, Title) Then Exit Sub

5         For Each c In Selection.Columns(1).Cells
6             If c.HasFormula Or (VarType(c.Value) = vbDouble Or VarType(c.Value) = vbError) Then
7                 Prompt = "Cannot justify cells containing numbers, formulas or errors."
8                 MsgBoxPlus Prompt, vbExclamation, Title
9                 Exit Sub
10            End If
11        Next c

12        Set CopyFrom = Selection.Columns(1)
13        With CopyFrom
14            If IsEmpty(.Cells(.Rows.Count, 1)) Then
15                Set BottomCell = .Cells(.Rows.Count, 1).End(xlUp)
16                If BottomCell.row < .row Or (BottomCell.row = 1 And IsEmpty(BottomCell.Value)) Then
17                    Throw "There is no text to justify", True
18                End If
19                Set CopyFrom = Range(.Cells(1, 1), BottomCell)
20            End If
21        End With

22        CurrentText = CopyFrom.Value
23        FontName = Selection.Cells(1, 1).Font.Name
24        If IsNull(FontName) Then
25            FontName = Selection.Cells(1, 1).Characters(1, 1).Font.Name
26        End If
27        FontSize = Selection.Cells(1, 1).Font.Size
28        If IsNull(FontSize) Then
29            FontSize = Selection.Cells(1, 1).Characters(1, 1).Font.Size
30        End If

31        AvailableWidth = Selection.Width
32        JustifiedText = sJustifyText(CurrentText, CStr(FontName), CInt(FontSize), AvailableWidth)
33        If VarType(JustifiedText) = vbString Then
34            If Left$(JustifiedText, 13) = "#sJustifyText" Then
35                If Right$(JustifiedText, 1) = "!" Then
36                    Throw JustifiedText
37                End If
38            End If
39        End If

40        If Selection.row + sNRows(JustifiedText) - 1 > ActiveSheet.Rows.Count Then
41            Prompt = "You cannot paste beyond the last row of the worksheet"
42            MsgBoxPlus Prompt, vbExclamation, Title
43            Exit Sub
44        End If

45        Set TargetRange = Selection.Resize(sNRows(JustifiedText), 1)
46        Prompt = TestForBlockingArrayFormulas(TargetRange)
47        If Len(Prompt) > 2 Then
48            MsgBoxPlus Prompt, vbExclamation, Title
49            Exit Sub
50        End If

51        Set OldSelection = Selection

52        If OldSelection.Rows.Count < TargetRange.Rows.Count Then
              Dim RangeToTest As Range
              Dim ShowPrompt As Boolean
53            Set RangeToTest = OldSelection.Offset(OldSelection.Rows.Count).Resize(TargetRange.Rows.Count - OldSelection.Rows.Count, 1)
54            Debug.Print RangeToTest.address
55            For Each c In RangeToTest.Cells
56                If Not IsEmpty(c.Value) Then
57                    ShowPrompt = True
58                    Exit For
59                End If
60            Next c
61        End If

62        If ShowPrompt Then
63            Prompt = "Overwrite these cells?" + vbLf + vbLf + "(Ctrl Z to undo)"
64            Application.GoTo TargetRange
65            If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel, "Justify Text") <> vbOK Then Exit Sub
66        End If
67        Set SUH = CreateScreenUpdateHandler()
68        BackUpRange Application.Union(OldSelection, TargetRange), shUndo
69        OldSelection.Columns(1).ClearContents
70        MyPaste TargetRange, JustifiedText
71        With TargetRange
72            .Font.Name = FontName
73            .Font.Size = FontSize
74        End With
75        Application.OnUndo "Undo Justify Text", "RestoreRange"
76        Exit Sub
ErrHandler:
77        SomethingWentWrong "#JustifyText (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestForBlockingArrayFormulas
' Author    : Philip Swannell
' Date      : 09-Oct-2016
' Purpose   : Sub-routine of JustifyText - checks if there are array formulas that will
'             prevent pasting in of the justified text. Return is "OK" or a helpful message to post.
' -----------------------------------------------------------------------------------------------------------------------
Private Function TestForBlockingArrayFormulas(RangeToTest As Range)
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
15                        If Application.Intersect(c.CurrentArray, RangeToTest).Cells.CountLarge <> c.CurrentArray.Cells.CountLarge Then
16                            TestForBlockingArrayFormulas = "You cannot change part of the array at " & AddressND(c.CurrentArray)
17                            Exit Function
18                        End If
19                    End If
20                Next c
21            End If
22        End If
23        TestForBlockingArrayFormulas = "OK"
24        Exit Function
ErrHandler:
25        Throw "#TestForBlockingArrayFormulas (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AbbreviateForCommandBar
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Command bar (as employed by ShowCommandBarPopup) cannot display text that
'             is wider than 407.25 points in the font used i.e. Segoe UI 9 point. This routine
'             abbreviates a file name or other string so that it will "fit". We take a few characters from
'             the start plus "..." plus as many characters from the right as we can.
' -----------------------------------------------------------------------------------------------------------------------
Public Function AbbreviateForCommandBar(ByVal LongFormText As String, Optional IsLitteral As Boolean) As String
          Const FontName = "Segoe UI"
          Const FontSize = 9
          Const MaxWidth = 407.25
          Dim Result As String
          Const NumCharsLeftPart = 10
          Dim NumChars As Long
          Dim NumCharsInRightPart As Long
          Dim PartialSum As Double

1         On Error GoTo ErrHandler

          'Because command bar caption cannot contain line feed character
2         If InStr(LongFormText, vbLf) > 0 Then
3             LongFormText = Replace(LongFormText, vbLf, "")
4         End If

5         If sStringWidth(LongFormText, FontName, FontSize)(1, 1) <= MaxWidth Then
6             AbbreviateForCommandBar = LongFormText
7         Else
              Dim AvailableWidth As Variant
              Dim CharacterArray As Variant
              Dim i As Long
              Dim WidthArray As Variant
8             NumChars = Len(LongFormText)
9             CharacterArray = sReshape(0, NumChars, 1)
10            For i = 1 To NumChars
11                CharacterArray(i, 1) = Mid$(LongFormText, i, 1)
12            Next i
13            WidthArray = sStringWidth(CharacterArray, FontName, FontSize)
14            Result = Left$(LongFormText, NumCharsLeftPart) + "..."
15            AvailableWidth = MaxWidth - sStringWidth(Result, FontName, FontSize)(1, 1)
16            PartialSum = WidthArray(NumChars, 1)
17            For i = NumChars - 1 To (NumCharsLeftPart + 3) Step -1
18                PartialSum = PartialSum + WidthArray(i, 1)
19                If PartialSum > AvailableWidth Then
20                    NumCharsInRightPart = NumChars - i - 1
21                    Exit For
22                End If
23            Next i
24            If NumCharsInRightPart = 0 Then NumCharsInRightPart = NumChars - NumCharsLeftPart - 3
25            AbbreviateForCommandBar = Result + Right$(LongFormText, NumCharsInRightPart)
26        End If
27        If IsLitteral Then        ' have to escape ampersands, not sure how to cope with litteral -- as first characters...
28            AbbreviateForCommandBar = Replace(AbbreviateForCommandBar, "&", "&&")
29        End If

30        Exit Function
ErrHandler:
31        AbbreviateForCommandBar = "#AbbreviateForCommandBar (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OpenTextFileNoRecalc
' Author     : Philip Swannell
' Date       : 27-Mar-2019
' Purpose    : Work in progress. Open a text file without the incredibly annoying recalculation of all open workbooks!
'              ToDo: Add to Ribbon and to BackStage if I can figure out how to do that.
'              But other approaches are to use SolumAddin function sFileShow or to use Excel ribbon > Data > From Text/CSV
' -----------------------------------------------------------------------------------------------------------------------
Function OpenTextFileNoRecalc(Optional FileName As String, _
        Optional RegKey As String = "TextFiles", _
        Optional FileFilter As String = "Text Files (*.txt; *.csv),*.txt;*.csv,All Files (*.*),*.*", _
        Optional Title As String = "Open Text file with no Excel recalculation", _
        Optional WithMRU As Boolean = True) As Workbook
          
          Dim wb As Excel.Workbook
          Dim ws As Worksheet
          Dim CalculationState() As Variant
          Dim i As Long
            
1         On Error GoTo ErrHandler
2         If FileName = vbNullString Then
3             FileName = GetOpenFilenameWrap(RegKey, FileFilter, , Title, , , WithMRU)
4             If FileName = "False" Then Exit Function
5         End If
6         If Not sFileExists(FileName) Then
7             Throw "Cannot find file '" + FileName + "'"
8         End If
            
          Dim NumSheets As Long
9         For Each wb In Application.Workbooks
10            NumSheets = NumSheets + wb.Worksheets.Count
11        Next

12        ReDim CalculationState(1 To NumSheets, 1 To 2)

          'Disable calculation!
13        i = 0
14        For Each wb In Application.Workbooks
15            For Each ws In wb.Worksheets
16                i = i + 1
17                Set CalculationState(i, 1) = ws
18                CalculationState(i, 2) = ws.EnableCalculation
19                ws.EnableCalculation = False
20            Next ws
21        Next

22        On Error Resume Next
23        Workbooks.OpenText FileName, Local:=True
24        If Err.Number <> 0 Then
25            MsgBoxPlus Err.Description, vbOKOnly + vbExclamation, "Open Text File"
26        Else
27            Set OpenTextFileNoRecalc = ActiveWorkbook
28        End If
29        On Error GoTo ErrHandler

          'Revert calculation status
30        For i = 1 To NumSheets
31            CalculationState(i, 1).EnableCalculation = CalculationState(i, 2)
32        Next

33        Exit Function
ErrHandler:
34        Throw "#OpenTextFileNoRecalc (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DontUseMeWithoutSpill
' Author     : Philip Swannell
' Date       : 10-Mar-2020
' Purpose    : For use in Workbook_Open of workbooks that require dynamic arrays.
' Parameters :
'  Wb:
' -----------------------------------------------------------------------------------------------------------------------
Sub DontUseMeWithoutSpill(wb As Workbook)
          Dim Prompt  As String
          Const Title = "Dynamic arrays formulas are not available"
1         If Not ExcelSupportsSpill Then
2             Prompt = "This Excel does not support dynamic array formulas, but the workbook '" + wb.Name + _
                  "' was written for Excel with dynamic arrays." + vbLf + vbLf + _
                  "Be careful, the formulas may give the wrong result in this version of Excel!"
3             MsgBoxPlus Prompt, vbOKOnly + vbExclamation, Title
4         End If
End Sub

