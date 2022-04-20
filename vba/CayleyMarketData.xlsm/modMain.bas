Attribute VB_Name = "modMain"
Option Explicit

Sub AmendForLiborTransition(ws As Worksheet)
          Dim SPH As clsSheetProtectionHandler
          Dim Target As Range
1         Set SPH = CreateSheetProtectionHandler(ws)
2         Set Target = RangeFromSheet(ws, "SwaptionVolParameters").Cells(8, 1)
3         With Target
4             .ClearFormats
5             .Value = "Libor Transition"
6         End With

7         With Target.Offset(1)
8             .ClearFormats
9             .Value = "FloatingLegType"
10        End With

11        With Target.Offset(1, 1)
12            .Clear
13            .Value = IIf(ws.Name = "EUR", "IBOR", "RFR")

14            With .Validation
15                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="RFR,IBOR"
16            End With

17        End With





End Sub


' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearCommentsFromActiveSheet
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Formats the active sheet, including removing comments added by Bloomberg feed.
'             Called from ShowMenu
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearCommentsFromActiveSheet()
1         On Error GoTo ErrHandler
2         If ActiveSheet.Parent Is ThisWorkbook Then
3             If IsCurrencySheet(ActiveSheet) Then
4                 FormatCurrencySheet ActiveSheet, True, Empty
5             ElseIf IsInflationSheet(ActiveSheet) Then
6                 FormatInflationSheet ActiveSheet, True, Empty
7             ElseIf ActiveSheet Is shFx Then
8                 FormatFxVolSheet True
9             End If
10        End If
11        Exit Sub
ErrHandler:
12        Throw "#ClearCommentsFromActiveSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearCommentsFromAllSheets
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Formats all the currency sheets and the FXVol sheet, including removing cell
'             comments added by the Bloomberg feeds. Called from ShowMenu.
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearCommentsFromAllSheets()
1         On Error GoTo ErrHandler
          Dim ws As Worksheet
2         For Each ws In ThisWorkbook.Worksheets
3             If IsCurrencySheet(ws) Then
4                 FormatCurrencySheet ws, True, Empty
5             ElseIf IsInflationSheet(ws) Then
6                 FormatInflationSheet ws, True, Empty
7             End If
8         Next
9         FormatFxVolSheet True
10        Exit Sub
ErrHandler:
11        Throw "#ClearCommentsFromAllSheets (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReleaseCleanup
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : Put the workbook in a tidy state for release.
' -----------------------------------------------------------------------------------------------------------------------
Function ReleaseCleanup()

          Const Title = "Release Cayley Market Data Workbook"
          Dim ChangesMade As Variant
          Dim ClearComments As Boolean
          Dim i As Long
          Dim j As Long
          Dim MsgBoxRes As VbMsgBoxResult
          Dim Prompt As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()

3         MsgBoxRes = MsgBoxPlus("Maintain Bloomberg feed comments or clear them out?", vbYesNoCancel + vbQuestion + vbDefaultButton1, Title, "Maintain (Recommended)", "C&lear out", , 90)
4         If MsgBoxRes = vbCancel Then Throw "Release aborted", True
5         ClearComments = (MsgBoxRes = vbNo)

6         Set SPH = CreateSheetProtectionHandler(shConfig)

7         ClearHiddenSheet
8         SortSheets
9         FormatFxVolSheet ClearComments
          
10        ChangesMade = sArrayRange("Worksheet", "Parameter", "Changed From", "Changed To")
11        SetCellForRelease shConfig, "Numeraire", "EUR", ChangesMade
12        SetCellForRelease shConfig, "CollateralCcy", "EUR", ChangesMade
13        SetCellForRelease shConfig, "HWRevert", 0.03, ChangesMade
14        SetCellForRelease shConfig, "SigmaStep", 1, ChangesMade
15        SetCellForRelease shConfig, "TStar", 15, ChangesMade
16        SetCellForRelease shConfig, "SCRiPTWorkbook", "SCRiPT2022.xlsm", ChangesMade
17        SetCellForRelease shConfig, "MarketDataFile", "..\data\market\20220228_solum.out", ChangesMade

18        If sNRows(ChangesMade) > 1 Then
19            For i = 2 To sNRows(ChangesMade)
20                For j = 3 To 4
21                    Select Case VarType(ChangesMade(i, j))
                          Case vbBoolean
22                            ChangesMade(i, j) = UCase(ChangesMade(i, j))
23                        Case vbString
24                            ChangesMade(i, j) = "'" & ChangesMade(i, j) & "'"
25                        Case Else
26                            ChangesMade(i, j) = CStr(ChangesMade(i, j))
27                    End Select
28                Next
29            Next
30            Prompt = "Release cleanup made the following changes to the workbook:" & vbLf & vbLf & _
                  sConcatenateStrings(sJustifyArrayOfStrings(ChangesMade, "SegoeUI", 9, vbTab), vbLf)
31            MsgBoxPlus Prompt, vbInformation, , , , , , 1000, , , 60, vbOK
32        End If

          'Activate top-left of each visible sheet
34        For Each ws In ThisWorkbook.Worksheets
35            ws.Protect , True, True
36            ws.Calculate
37            If IsCurrencySheet(ws) Then
38                FormatCurrencySheet ws, ClearComments, True
39            ElseIf IsInflationSheet(ws) Then
40                FormatInflationSheet ws, ClearComments, True
41            End If

42            If ws.CodeName <> "sh" & Replace(ws.Name, " ", "_") Then
43                ThisWorkbook.VBProject.VBComponents(ws.CodeName).Name = "shTempName"
44                ThisWorkbook.VBProject.VBComponents(ws.CodeName).Name = "sh" & Replace(ws.Name, " ", "_")
45            End If

46            If ws.Visible = xlSheetVisible Then
47                Application.GoTo ws.Cells(1, 1)
48                ActiveWindow.DisplayGridlines = False
49                ActiveWindow.DisplayHeadings = False
50            End If
51        Next ws
52        HideUnhideSheets ThisWorkbook, sArrayStack("EUR", "CAD", "USD", "GBP", "CHF", "JPY"), shConfig.Range("Numeraire"), "All"
53        With shFx.Range("TheFilters")
54            .ClearContents
55            .Cells(1, 1).Value = shConfig.Range("Numeraire")
56        End With
57        Application.GoTo shFx.Cells(1, 1)

58        Exit Function
ErrHandler:
          'Must return an error string since we call via Application.Run.
59        ReleaseCleanup = "#ReleaseCleanup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetCellForRelease(ws As Worksheet, RangeName As String, NewValue As Variant, ByRef ChangesMade)
          Dim c As Range
1         Set c = RangeFromSheet(ws, RangeName)
2         If Not sNearlyequals(c.Value, NewValue) Then
3             ChangesMade = sArrayStack(ChangesMade, sArrayRange(ws.Name, RangeName, c.Value, NewValue))
4             SafeSetCellValue c, NewValue
5         End If
End Function



' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortSheets
' Author    : Philip Swannell
' Date      : 13-Jul-2016
' Purpose   : Put the sheets of this workbook into the preferred order
' -----------------------------------------------------------------------------------------------------------------------
Sub SortSheets()
          Dim i As Long
          Dim sheetList
          Dim SheetListCcy As Variant
          Dim SheetListInflation As Variant
          Dim STK_Ccy As clsStacker
          Dim STK_Inflation As clsStacker
          Dim SUH As clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set STK_Ccy = CreateStacker(): Set STK_Inflation = CreateStacker()
4         For Each ws In ThisWorkbook.Worksheets
5             If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
6             If IsCurrencySheet(ws) Then
7                 STK_Ccy.StackData ws.Name
8             ElseIf IsInflationSheet(ws) Then
9                 STK_Inflation.StackData ws.Name
10            End If
11        Next
12        SheetListCcy = sSortedArray(STK_Ccy.report)
13        SheetListInflation = sSortedArray(STK_Inflation.report)

14        sheetList = sArrayStack(shFx.Name, shCredit.Name, shHiddenSheet.Name, SheetListCcy, SheetListInflation, _
                                  shHistoricalCorrEUR.Name, shHistoricalCorrUSD.Name, shHistoricalCorrGBP.Name, shConfig.Name, shAudit.Name)

15        If Not (ThisWorkbook.Worksheets(1) Is shFx) Then
16            shFx.Move Before:=ThisWorkbook.Worksheets(1)
17        End If

18        For i = 2 To sNRows(sheetList)
19            If ThisWorkbook.Worksheets(i).Name <> sheetList(i, 1) Then
20                ThisWorkbook.Worksheets(sheetList(i, 1)).Move After:=ThisWorkbook.Worksheets(i - 1)
21            End If
22        Next i

23        shHiddenSheet.Visible = xlSheetHidden
24        shStaticData.Visible = xlSheetHidden

25        Exit Sub
ErrHandler:
26        Throw "#SortSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnprotectAllSheets
' Author    : Philip Swannell
' Date      : 05-Nov-2015
' Purpose   : Unprotects all sheets, attached to button on Audit sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub UnprotectAllSheets()
          Dim oldVisState
          Dim origSheet As Worksheet
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set origSheet = ActiveSheet
4         For Each ws In ThisWorkbook.Worksheets
5             ws.Unprotect
6             oldVisState = ws.Visible
7             ws.Visible = xlSheetVisible
8             GroupingButtonDoAllOnSheet ws, True
9             ws.Activate
10            ActiveWindow.DisplayHeadings = True
11            ws.Visible = oldVisState

12        Next
13        origSheet.Activate
14        Exit Sub
ErrHandler:
15        Throw "#UnprotectAllSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnhideAllCurrencySheets
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Called from button on Audit sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub UnhideAllCurrencySheets()

1         On Error GoTo ErrHandler
2         HideUnhideSheets ThisWorkbook, "all", "EUR", "all"
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#UnhideAllCurrencySheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HideUnhideSheets
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Set Visible property of sheets in the workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub HideUnhideSheets(MarketWb As Workbook, CurrenciesToShow As Variant, ByVal Numeraire As String, InflationIndicesToShow)
          Dim ShowAllCcys As Boolean
          Dim ShowAllII As Boolean
          Dim SUH As clsScreenUpdateHandler
          Dim viz As XlSheetVisibility
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
          Dim AnyChanges As Boolean
          Dim exSH As clsexcelStateHandler
          Dim origActiveSheet As Worksheet

2         Set exSH = CreateexcelStateHandler(, , , , , True)
3         Set SUH = CreateScreenUpdateHandler()

4         Set origActiveSheet = MarketWb.ActiveSheet
5         ShowAllCcys = sArraysIdentical(CurrenciesToShow, "All")
6         ShowAllII = sArraysIdentical(InflationIndicesToShow, "All")

7         For Each ws In MarketWb.Worksheets
8             viz = xlSheetHidden
9             If IsCurrencySheet(ws) Then
10                If ShowAllCcys Then
11                    viz = xlSheetVisible
12                ElseIf IsNumber(sMatch(ws.Name, CurrenciesToShow)) Then
13                    viz = xlSheetVisible
14                ElseIf ws.Name = Numeraire Then
15                    viz = xlSheetVisible
16                Else
17                    viz = xlSheetHidden
18                End If
19            ElseIf IsInflationSheet(ws) Then
20                If ShowAllII Then
21                    viz = xlSheetVisible
22                ElseIf IsNumber(sMatch(ws.Name, InflationIndicesToShow)) Then
23                    viz = xlSheetVisible
24                Else
25                    viz = xlSheetHidden
26                End If
27            Else
28                If ws Is shFx Or ws Is shCredit Or ws Is shConfig Or ws Is shAudit Then
29                    viz = xlSheetVisible
30                ElseIf ws Is shHiddenSheet Or ws Is shStaticData Then
31                    viz = xlSheetHidden
32                ElseIf Left(ws.Name, 14) = "HistoricalCorr" Then
33                    If Right(ws.Name, 3) = Numeraire Or ShowAllCcys = True Then
34                        viz = xlSheetVisible
35                    Else
36                        viz = xlSheetHidden
37                    End If
38                End If
39            End If
40            If ws.Visible <> viz Then
41                AnyChanges = True
42                ws.Visible = viz
43            End If
44        Next ws

45        If AnyChanges Then
46            If origActiveSheet.Visible Then
47                origActiveSheet.Activate
48            End If
49        End If

50        Set exSH = Nothing
51        Set SUH = Nothing

52        Exit Sub
ErrHandler:
53        Throw "#HideUnhideSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ForceAllCorrelationsToBeSymmetric
' Author    : Philip Swannell
' Date      : 03-May-2017
' Purpose   : Ad-hoc method to put the "make me symmetric" formulas back in to the correlation sheets
' -----------------------------------------------------------------------------------------------------------------------
Sub ForceAllCorrelationsToBeSymmetric()
          Dim ws As Worksheet
1         For Each ws In ThisWorkbook.Worksheets
2             If Left(ws.Name, 14) = "HistoricalCorr" Then
3                 ForceCorrelationsToBeSymmetric ws
4             End If
5         Next
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ForceCorrelationsToBeSymmetric
' Author    : Philip Swannell
' Date      : 21-Jun-2016
' Purpose   : enters array formulas so that each column beneath the diagonal is set to an
'             array formula pointing to the corresponding row above the diagonal
' -----------------------------------------------------------------------------------------------------------------------
Sub ForceCorrelationsToBeSymmetric(ws As Worksheet)
          Dim CorrRng As Range
          Dim i As Long
          Dim N As Long
          Dim SourceRange As Range
          Dim SPH As clsSheetProtectionHandler
          Dim TargetRange As Range

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(ws)

3         Set CorrRng = sexpandRightDown(RangeFromSheet(ws, "HistCorrMatrix"))
4         With CorrRng
5             Set CorrRng = .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1)

6             N = .Columns.Count

7             For i = 1 To N - 1
8                 Set SourceRange = Range(.Cells(i, i + 1), .Cells(i, N))
9                 Set TargetRange = Range(.Cells(i + 1, i), .Cells(N, i))
10                TargetRange.FormulaArray = "=TRANSPOSe(" + Replace(SourceRange.Address, "$", "") + ")"
11            Next i
12        End With

13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#ForceCorrelationsToBeSymmetric (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ImportHistoricalCorr
' Author    : Philip Swannell
' Date      : 21-Jun-2016
' Purpose   : Grabs data from another sheet, ideally a sheet created by Correlation Matrix Generator,
'             and pastes it to the HistoricalCorr sheet.
' -----------------------------------------------------------------------------------------------------------------------
Sub ImportHistoricalCorr(TargetSheet As Worksheet)

          Dim AllPossibleAddresses As Variant
          Dim AllPossibleRangeNames As Variant
          Dim Default As String
          Dim i As Long
          Dim j As Long
          Dim N As Long
          Dim Notes As String
          Dim R As Range
          Dim SUH As clsScreenUpdateHandler
          Dim TheChoices As Variant
          Dim TheMatrix As Variant
          Dim TopText As String
          Dim wb As Workbook
          Dim ws As Worksheet
          Const SomeOtherRange = "<Some other range>"
          Dim RangeName As String
          Dim RequiredNumeraire

1         On Error GoTo ErrHandler

2         RequiredNumeraire = Right(TargetSheet.Name, 3)

3         AllPossibleAddresses = CreateMissing()
4         AllPossibleRangeNames = CreateMissing()

5         For Each wb In Application.Workbooks
6             For Each ws In wb.Worksheets
7                 If ws.Name = RequiredNumeraire Then
8                     For i = 1 To 3
9                         RangeName = Choose(i, "SampleCorrelation", "SampleCorrelation2", "RobustCorrelation")
10                        If IsInCollection(ws.Names, RangeName) Then
11                            Set R = Nothing
12                            On Error Resume Next
13                            Set R = ws.Names(RangeName).RefersToRange
14                            On Error GoTo ErrHandler
15                            If Not R Is Nothing Then
16                                Set R = R.Offset(-1, -1).Resize(R.Rows.Count + 1, R.Columns.Count + 1)
17                                AllPossibleAddresses = sArrayStack(AllPossibleAddresses, R.Address(external:=True))
18                                AllPossibleRangeNames = sArrayStack(AllPossibleRangeNames, RangeName + " " + vbTab + CStr(R.Rows.Count - 1) + "x" + CStr(R.Columns.Count - 1))
19                            End If
20                        End If
21                    Next i
22                End If
23            Next ws
24        Next wb
endLoop:

25        If IsMissing(AllPossibleAddresses) Then
26            Default = ""
27        ElseIf sNRows(AllPossibleAddresses) = 1 Then
28            Default = AllPossibleAddresses(1, 1)
29        Else
30            AllPossibleAddresses = sArrayStack(AllPossibleAddresses, SomeOtherRange)
31            AllPossibleRangeNames = sArrayStack(AllPossibleRangeNames, "")
32            TopText = "There is more than one correlation matrix available." + vbLf + "Please select which one you want to import."
33            TheChoices = sJustifyArrayOfStrings(sArrayRange(AllPossibleRangeNames, AllPossibleAddresses), "Tahoma", 9, vbTab)
34            Default = ShowSingleChoiceDialog(TheChoices, , , , , , TopText)
35            If Default = "" Then Exit Sub
36            Default = sVLookup(Default, sArrayRange(TheChoices, AllPossibleAddresses))
37            If Default = SomeOtherRange Then Default = ""
38        End If

39        Set R = Nothing
          Dim Prompt As String

          'If Not Default = "" Then
          '    Application.Goto Range(Default)
          'end If

40        If Default = "" Then
41            Prompt = "Please select a range containing the correlation matrix with headers" + vbLf + vbLf + "Use the workbook 'c:\SolumWorkbooks\Correlation Matrix Generator.xlsm' to generate the data."
42        Else
43            Prompt = "Please confirm this is the range containing the correlation matrix with headers, or select another range."
44        End If

45        If Default = "" Then
46            On Error Resume Next
47            Set R = Application.InputBox(Prompt, , Default, , , , , 8)
48            On Error GoTo ErrHandler
49            If R Is Nothing Then Exit Sub
50        Else
51            Set R = Range(Default)
52        End If

53        Set SUH = CreateScreenUpdateHandler()

          'Sanity check
54        If R.Columns.Count <> R.Rows.Count Then Throw "Range must have the same number of rows as columns"
55        For i = 2 To R.Columns.Count
56            Select Case Right(CStr(R.Cells(1, i).Value), 3)
              Case " IR", " FX"
                  'OK
57            Case Else
58                Throw "Labels in the top row must end with either ' FX' or 'IR', but cell " + R.Cells(1, i).Address(external:=True) + " does not"
59            End Select
60            If CStr(R.Cells(i, 1)) <> CStr(R.Cells(1, i)) Then
61                Throw "Labels in the first column must be the same as labels in the top row but cells " + R.Cells(1, i).Address(external:=True) & " and " + R.Cells(i, 1).Address(external:=True) + " are different"
62            End If
63        Next i

64        TheMatrix = R.Offset(1, 1).Resize(R.Rows.Count - 1, R.Columns.Count - 1)
65        N = sNRows(TheMatrix)

66        For i = 1 To N
67            For j = 1 To N
68                If Not IsNumber(TheMatrix(i, j)) Then Throw "Non numbers detected in the selected range, e.g. at " + R.Cells(i + 1, j + 1).Address(external:=True)
69            Next j
70        Next i

71        For i = 1 To N
72            If TheMatrix(i, i) <> 1 Then Throw "On-diagonal cells must be 1, but cell " + R.Cells(i + 1, i + 1).Address(external:=True) + " is not"
73        Next i

74        For i = 1 To N
75            For j = 1 To i - 1
76                If TheMatrix(i, j) <> TheMatrix(j, i) Then Throw "Selected range is not a symmetric matrix with headers, for example " + R.Cells(i + 1, j + 1).Address(external:=True) & " and " + R.Cells(j + 1, i + 1).Address(external:=True) + " are different"
77                If TheMatrix(i, j) > 1 Or TheMatrix(i, j) < -1 Then Throw "All cells in the selected range (excluding headers) must be in the range -1 to 1, but " + Cells(i + 1, j + 1).Address(external:=True) + " is not"
78            Next j
79        Next i

80        If IsInCollection(R.Parent.Names, "StartDate") Then
81            If IsInCollection(R.Parent.Names, "endDate") Then
82                On Error Resume Next
83                Notes = "Data imported from range " + RelevantName(R) + " of workbook " + R.Parent.Parent.FullName + " which used data from " + Format(R.Parent.Range("StartDate").Value, "d-mmm-yyyy") + " to " + Format(R.Parent.Range("endDate").Value, "d-mmm-yyyy")
84                On Error GoTo ErrHandler
85            End If
86        End If

          Dim DataToImport As Variant
          Dim NewRange As Range
          Dim NewRangeNumbers As Range
          Dim OldRange As Range
          Dim SPH As clsSheetProtectionHandler

87        Set SPH = CreateSheetProtectionHandler(TargetSheet)
88        DataToImport = R.Value2

89        TargetSheet.Range("B2").Value = Notes

90        Set OldRange = sexpandRightDown(TargetSheet.Range("HistCorrMatrix"))
91        OldRange.Clear

92        Set NewRange = TargetSheet.Range("HistCorrMatrix").Resize(R.Rows.Count, R.Columns.Count)
93        Set NewRangeNumbers = NewRange.Offset(1, 1).Resize(NewRange.Rows.Count - 1, NewRange.Columns.Count - 1)

94        NewRange.Resize(NewRange.Rows.Count + 1000, NewRange.Columns.Count + 1000).Clear
95        NewRange.Value = DataToImport
96        NewRange.HorizontalAlignment = xlHAlignCenter
97        NewRangeNumbers.NumberFormat = "0%"
98        AddGreyBorders NewRangeNumbers, True
99        DoConditionalFormatting NewRangeNumbers
          ' ForceCorrelationsToBeSymmetric TargetSheet

100       Application.GoTo TargetSheet.Cells(1, 1)
101       TargetSheet.Activate
102       Application.ScreenUpdating = False
103       Application.ScreenUpdating = True

104       Exit Sub
ErrHandler:
105       SomethingWentWrong "#ImportHistoricalCorr (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RelevantName
' Author    : Philip
' Date      : 22-Jun-2017
' Purpose   : sub-routine of ImportHistoricalCorr
' -----------------------------------------------------------------------------------------------------------------------
Private Function RelevantName(R As Range)
          Dim Q As Range
1         On Error GoTo ErrHandler
2         Set Q = R.Offset(1, 1).Resize(R.Rows.Count - 1, R.Columns.Count - 1)
3         RelevantName = sStringBetweenStrings(Q.Name.Name, "!")
4         Exit Function
ErrHandler:
5         RelevantName = Replace(R.Address, "$", "")
End Function

Sub DoConditionalFormatting(R As Range)

1         On Error GoTo ErrHandler
2         With R
3             .FormatConditions.AddColorScale ColorScaleType:=2
4             .FormatConditions(.FormatConditions.Count).SetFirstPriority
5             .FormatConditions(1).ColorScaleCriteria(1).Type = _
              xlConditionValueNumber
6             .FormatConditions(1).ColorScaleCriteria(1).Value = -1
7             With .FormatConditions(1).ColorScaleCriteria(1).FormatColor
8                 .Color = 2650623
9                 .TintAndShade = 0
10            End With
11            .FormatConditions(1).ColorScaleCriteria(2).Type = _
              xlConditionValueNumber
12            .FormatConditions(1).ColorScaleCriteria(2).Value = 1
13            With .FormatConditions(1).ColorScaleCriteria(2).FormatColor
14                .Color = 10285055
15                .TintAndShade = 0
16            End With
17        End With

18        Exit Sub
ErrHandler:
19        Throw "#DoConditionalFormatting (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


