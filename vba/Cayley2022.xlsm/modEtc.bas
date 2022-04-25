Attribute VB_Name = "modEtc"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunThisAtTopOfCallStack
' Author     : Philip Swannell
' Date       : 16-Mar-2022
' Purpose    : When VBA code execution starts, we need to ensure that EnableEvents etc are set as they should be
'              Hopefully, the only circumstance in which the properties are likely to have the wrong values are during code development
'              with debugging activity causing code to exit not cleanly. But calling this method at the top of the call stack is certainly safe.
' -----------------------------------------------------------------------------------------------------------------------
Sub RunThisAtTopOfCallStack()

1         On Error GoTo ErrHandler
2         If Application.Cursor <> xlDefault Then
3             Application.Cursor = xlDefault
4         End If
5         If Not Application.EnableEvents Then
6             Application.EnableEvents = True
7         End If
8         SetCalculationToManual
9         Application.StatusBar = False
10        CheckAddinVersion "SolumAddin.xlam", gMinimumSolumAddinVersion
11        CheckAddinVersion "SolumSCRiPTUtils.xlam", gMinimumSolumSCRiPTUtilsVersion

12        Exit Sub
ErrHandler:
13        Throw "#RunThisAtTopOfCallStack (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromConfig
' Author     : Philip Swannell
' Date       : 14-Mar-2022
' Purpose    : For better relocatability, files and folders on the Config sheet may be given as paths relative to
'              the folder of this workbook. This function returns a file's full path.
' -----------------------------------------------------------------------------------------------------------------------
Function FileFromConfig(NameOnConfig As String)

          Dim Res As String
1         On Error GoTo ErrHandler
2         Res = RangeFromSheet(shConfig, NameOnConfig, False, True, False, False, False).Value
3         FileFromConfig = sJoinPath(ThisWorkbook.Path, Res)

4         Exit Function
ErrHandler:
5         Throw "#FileFromConfig (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DictAdd
' Author    : Philip Swannell
' Date      : 04-Oct-2016
' Purpose   : Adds item to dictionary, but (unlike native .Add method) if the item already
'             exists then we overwrite
' -----------------------------------------------------------------------------------------------------------------------
Function DictAdd(D As Dictionary, key As String, Item As Variant)
1         On Error GoTo ErrHandler
2         If D.Exists(key) Then D.Remove key
3         D.Add key, Item
4         Exit Function
ErrHandler:
5         Throw "#DictAdd (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DictGet
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : Gets an item from a dictionary but rather than returning Empty if the
'             Key does not exist we throw an error.
' -----------------------------------------------------------------------------------------------------------------------
Function DictGet(D As Dictionary, key As String)
1         On Error GoTo ErrHandler
2         If Not D.Exists(key) Then
3             Throw "Dictionary does not contain item '" & key & "'"
4         End If
5         If VarType(D.Item(key)) = vbObject Then
6             Set DictGet = D.Item(key)
7         Else
8             DictGet = D.Item(key)
9         End If

10        Exit Function
ErrHandler:
11        Throw "#DictGet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetItem
' Author     : Philip Swannell
' Date       : 27-Jan-2022
' Purpose    : Mmm similar function to DictGet, but different signature. ToDo - clean up
' -----------------------------------------------------------------------------------------------------------------------
Function GetItem(D As Variant, key As String)
1         On Error GoTo ErrHandler
2         If D.Exists(key) Then
3             If IsObject(D(key)) Then
4                 Set GetItem = D(key)
5             Else
6                 GetItem = D(key)
7             End If
8         Else
9             Throw "key '" & key & "' not found in dictionary"
10        End If

11        Exit Function
ErrHandler:
12        Throw "#GetItem (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddRangeNames
' Author    : Philip Swannell
' Date      : 12-May-2015
' Purpose   : Quick tool to set up range names on the market data sheets
' -----------------------------------------------------------------------------------------------------------------------
Sub AddRangeNames()
          Dim c As Range
          Dim TheRange As Range

1         For Each c In Selection.Cells
2             If Not IsEmpty(c.Value) Then
3                 Set TheRange = c.CurrentRegion
4                 With TheRange
5                     Set TheRange = .offset(1).Resize(.Rows.Count - 1)
6                 End With
7                 ActiveSheet.Names.Add c.Value, TheRange
8                 AddGreyBorders TheRange, True
9             End If
10        Next
11    End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetSheetVisibility
' Author    : Philip Swannell
' Date      : 07-Oct-2016
' Purpose   : Set the visibility of each sheet in the workbook. Called from release cleanup
' -----------------------------------------------------------------------------------------------------------------------
Sub SetSheetVisibility(ShowAll As Boolean)
          Dim ws As Worksheet

1         ThisWorkbook.Unprotect
2         On Error GoTo ErrHandler
3         If ShowAll Then
4             For Each ws In ThisWorkbook.Worksheets
5                 ws.Visible = xlSheetVisible
6             Next ws
7         Else
8             For Each ws In ThisWorkbook.Worksheets
9                 Select Case ws.Name
                      Case shCreditUsage.Name, shTable.Name, shScenarioDefinition.Name, shScenarioResults.Name, _
                          shTradesViewer.Name, shBubbleChart.Name, shBarChart.Name, _
                          shConfig.Name, shAudit.Name, shHistoricalData.Name, shScenarioCompare.Name, shToDo.Name, shExportToTMS.Name
10                        ws.Visible = xlSheetVisible
11                    Case shLinesHistory.Name, shWhoHasLines.Name, shCommentEditor.Name, _
                          shFutureTrades.Name
12                        ws.Visible = xlSheetHidden
13                    Case Else
14                        ws.Visible = xlSheetHidden
15                End Select
16            Next ws
17        End If
18        ThisWorkbook.Protect
19        Exit Sub
ErrHandler:
20        Throw "#SetSheetVisibility (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReleaseCleanup
' Author    : Philip Swannell
' Date      : 29-May-2015
' Purpose   : Put workbook in good state for releasing
' -----------------------------------------------------------------------------------------------------------------------
Sub ReleaseCleanup()

          Dim ChangesMade As Variant
          Dim i As Long
          Dim j As Long
          Dim Res
          Dim SUH As Object
          Dim ws As Worksheet
          Const BankName = "WPAC_AU_SYD"        'Choose an "interesting" chart
          Dim globalsError As String

1         On Error GoTo ErrHandler

          Dim vNum
2         vNum = Application.Workbooks("SolumAddin.xlam").Worksheets("Audit").Range("Headers").Cells(2, 1).Value
3         If gMinimumSolumAddinVersion <> vNum Then globalsError = "VBA constant modGlobals.gMinimumSolumAddinVersion should be updated from " & gMinimumSolumAddinVersion & " to " & CStr(vNum)
4         vNum = Application.Workbooks("SolumSCRiPTUtils.xlam").Worksheets("Audit").Range("Headers").Cells(2, 1).Value
5         If gMinimumSolumSCRiPTUtilsVersion <> vNum Then globalsError = globalsError & IIf(globalsError = "", "", " and ") & "VBA constant modGlobals.gMinimumSolumSCRiPTUtilsVersion should be updated from " & gMinimumSolumSCRiPTUtilsVersion & " to " & CStr(vNum)
6         If globalsError <> "" Then Throw globalsError

          Dim Prompt As String
7         Prompt = "Do quick cleanup?"
8         Res = MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion + vbDefaultButton2, _
              ThisWorkbook.Name, "Yes, Quick", "No, Full", "Cancel, Abort Release")

9         If Res = vbCancel Then
10            Throw "Release aborted", True
11        ElseIf Res = vbYes Then
12            Application.GoTo shCreditUsage.Cells(1, 1)
13            Exit Sub
14        End If

15        Application.FormulaBarHeight = 1
16        Set SUH = CreateScreenUpdateHandler()

17        ThisWorkbook.Protect , False, False

18        SetSheetVisibility False

19        For Each ws In ThisWorkbook.Worksheets
20            ResizeCommentsOnSheet ws
21            Res = ws.UsedRange.Rows.Count
22            ws.Protect , True, True
23            If ws.Visible = xlSheetVisible Then
24                ActiveWindow.DisplayGridlines = False
25                ActiveWindow.DisplayHeadings = False
26                ActiveWindow.Zoom = 100
27                If ws.EnableSelection = xlNoRestrictions Then
28                    Application.GoTo ws.Cells(1, 1)
29                Else
30                    ActiveWindow.ScrollIntoView 0, 0, 1, 1
31                End If
32            End If
33        Next ws

34        ChangesMade = sArrayRange("Worksheet", "Parameter", "Changed From", "Changed To")

35        SetCellForRelease shCreditUsage, "Filter1Value", BankName, ChangesMade
36        SetCellForRelease shCreditUsage, "FilterBy1", "Counterparty Parent", ChangesMade
37        SetCellForRelease shCreditUsage, "FilterBy2", "None", ChangesMade
38        SetCellForRelease shCreditUsage, "Filter2Value", "None", ChangesMade
39        SetCellForRelease shCreditUsage, "IncludeExtraTrades", False, ChangesMade
40        SetCellForRelease shCreditUsage, "IncludeFutureTrades", False, ChangesMade
41        SetCellForRelease shCreditUsage, "IncludeAssetClasses", "Rates and Fx", ChangesMade
42        SetCellForRelease shCreditUsage, "PortfolioAgeing", 0, ChangesMade
43        SetCellForRelease shCreditUsage, "FxShock", 1, ChangesMade
44        SetCellForRelease shCreditUsage, "FxVolShock", 1, ChangesMade
45        SetCellForRelease shCreditUsage, "ModelType", MT_HW, ChangesMade
46        SetCellForRelease shCreditUsage, "NumMCPaths", 255, ChangesMade
47        SetCellForRelease shCreditUsage, "NumObservations", 100, ChangesMade
48        SetCellForRelease shCreditUsage, "TradesScaleFactor", 1, ChangesMade
49        SetCellForRelease shCreditUsage, "LinesScaleFactor", 1, ChangesMade
50        SetCellForRelease shCreditUsage, "ExtraTradesAre", "Fx Airbus sells USD, buys EUR", ChangesMade
51        SetCellForRelease shConfig, "FxTradesCSVFile", "..\data\trades\ExampleFxTrades.csv", ChangesMade
52        SetCellForRelease shConfig, "RatesTradesCSVFile", "..\data\trades\ExampleRatesTrades.csv", ChangesMade
53        SetCellForRelease shConfig, "AmortisationCSVFile", "..\data\trades\ExampleAmortisation.csv", ChangesMade
54        SetCellForRelease shConfig, "LinesWorkbook", "CayleyLines.xlsm", ChangesMade
55        SetCellForRelease shConfig, "MarketDataWorkbook", "CayleyMarketData.xlsm", ChangesMade
56        SetCellForRelease shConfig, "HedgeHorizon", 8, ChangesMade
57        SetCellForRelease shExportToTMS, "FeedRates", True, ChangesMade
58        SetCellForRelease shExportToTMS, "ExportTrades", True, ChangesMade
59        SetCellForRelease shExportToTMS, "ExportMarketData", True, ChangesMade
60        SetCellForRelease shExportToTMS, "ExportTable", True, ChangesMade
61        SetCellForRelease shExportToTMS, "ExportCharts", True, ChangesMade
62        SetCellForRelease shBubbleChart, "FxBreakEvenFloor", 0, ChangesMade
63        RangeFromSheet(shExportToTMS, "Scenarios").Value = sArraySquare(sReshape(True, 3, 1), _
              sArrayStack( _
              "C:\CayleyScenarios\$30_bn_pa_3,4,5Y_100%_fwds_Path_Jan-09-Jan-11.sdf", _
              "C:\CayleyScenarios\$30_bn_pa_3,4,5Y_100%_fwds_Path_Jan-07-Jan-09.sdf", _
              "C:\CayleyScenarios\$30_bn_pa_3,4,5Y_100%_fwds_Path_Jan-00-Jan-02.sdf"), _
              sReshape(False, 17, 1), sReshape(Empty, 17, 1))

64        If sNRows(ChangesMade) > 1 Then
65            For i = 2 To sNRows(ChangesMade)
66                For j = 3 To 4
67                    Select Case VarType(ChangesMade(i, j))
                          Case vbBoolean
68                            ChangesMade(i, j) = UCase(ChangesMade(i, j))
69                        Case vbString
70                            ChangesMade(i, j) = "'" & ChangesMade(i, j) & "'"
71                        Case Else
72                            ChangesMade(i, j) = CStr(ChangesMade(i, j))
73                    End Select
74                Next
75            Next
76            Prompt = "Release cleanup made the following changes to the workbook:" & vbLf & vbLf & _
                  sConcatenateStrings(sJustifyArrayOfStrings(ChangesMade, "SegoeUI", 9, vbTab), vbLf)
77            MsgBoxPlus Prompt, vbInformation, , , , , , 1000, , , 60, vbOK
78        End If

79        ResetTableButtons
80        AlignMenuButtons
81        ClearFutureTrades
82        ClearOutTradeValuesSheet
83        FormatExportToTMS

84        OpenOtherBooks
85        PrepareScenarioDefinitionSheetForRelease
86        PrepareForCalculation BankName, True, True, True
87        With shBarChart.Range("SortBy")
88            If .Value <> "THR 3Y" Then .Value = "THR 3Y"
89        End With
90        For Each ws In ThisWorkbook.Worksheets
91            ws.Calculate        'Updates BubbleChart, BarChart
92        Next ws
93        ThisWorkbook.Windows(1).Activate
94        Application.GoTo RangeFromSheet(shCreditUsage, "Filter1Value")
95        JuliaLaunchForCayley
96        RunCreditUsageSheet "Standard", True, True, True
97        FormatCreditUsageSheet True
98        ShowHidePFEData False
99        ThisWorkbook.Protect , True, True

100       Exit Sub
ErrHandler:
101       SomethingWentWrong "#ReleaseCleanup (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Function SetCellForRelease(ws As Worksheet, RangeName As String, NewValue As Variant, ByRef ChangesMade)
          Dim c As Range
1         Set c = RangeFromSheet(ws, RangeName)
2         If Not sNearlyEquals(c.Value, NewValue) Then
3             ChangesMade = sArrayStack(ChangesMade, sArrayRange(ws.Name, RangeName, c.Value, NewValue))
4             SafeSetCellValue c, NewValue
5         End If
End Function

Function SafeMax(a, b)
1         On Error GoTo ErrHandler
2         If a > b Then
3             SafeMax = a
4         Else
5             SafeMax = b
6         End If
7         Exit Function
ErrHandler:
8         SafeMax = "#" & Err.Description & "!"
End Function

Function SafeMin(a, b)
1         On Error GoTo ErrHandler
2         If a > b Then
3             SafeMin = b
4         Else
5             SafeMin = a
6         End If
7         Exit Function
ErrHandler:
8         SafeMin = "#" & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ChooseAnchorObject
' Author     : Philip Swannell
' Date       : 24-Feb-2022
' Purpose    : To assist with positioning the Command bar menus that appear via calls to ShowCommandBarPopup
' -----------------------------------------------------------------------------------------------------------------------
Function ChooseAnchorObject() As Object
1         On Error GoTo ErrHandler
2         If VarType(Application.Caller) = vbString Then
3             Set ChooseAnchorObject = Nothing
4         Else
5             If IsInCollection(ActiveSheet.Buttons, "butMenu") Then
6                 Set ChooseAnchorObject = ActiveSheet.Buttons("butMenu")
7                 Exit Function
8             End If
9         End If

10        Exit Function
ErrHandler:
11        Throw "#ChooseAnchorObject (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShiftF9Response
' Author     : Philip Swannell
' Date       : 02-Mar-2022
' Purpose    : Links ShiftF9 through to "recalculating" worksheets, which involves calling methods to refresh the data
'              on the worksheets.
' -----------------------------------------------------------------------------------------------------------------------
Sub ShiftF9Response()
1         On Error GoTo ErrHandler

2         If ActiveSheet Is shCreditUsage Then
3             If Not OtherBooksAreOpen(True, True, True) Then PleaseOpenOtherBooks
4             MenuCreditUsageSheet "Calculate"
5         ElseIf ActiveSheet Is shScenarioDefinition Then
6             If Not OtherBooksAreOpen(True, True, True) Then PleaseOpenOtherBooks
7             RefreshScenarioDefinition True, False
8         Else
9             SetKeys False 'If SetKeys is always being called when it needs to be we should never hit this line.
10            ActiveSheet.Calculate
11        End If
12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#ShiftF9Response (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub F8Response()

1         On Error GoTo ErrHandler
2         If ActiveSheet Is shCreditUsage Then
3             MenuCreditUsageSheet
4         ElseIf ActiveSheet Is shTable Then
5             MenuTableSheet
6         ElseIf ActiveSheet Is shScenarioDefinition Then
7             MenuScenarioDefinitionSheet
8         ElseIf ActiveSheet Is shScenarioResults Then
9             MenuScenarioResultsSheet
10        ElseIf ActiveSheet Is shBarChart Then
11            MenuBarChart
12        ElseIf ActiveSheet Is shHistoricalData Then
13            MenuHistoricData
14        ElseIf ActiveSheet Is shScenarioCompare Then
15            MenuScenarioCompare
16        ElseIf ActiveSheet Is shTradesViewer Then
17            MenuTradesViewerSheet
18        ElseIf ActiveSheet Is shExportToTMS Then
19            RunExport
20        Else
21            SetKeys False
22        End If

23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#F8Response (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub SetKeys(SwitchOn As Boolean)
          Const TheKey = "{F8}"
          Const TheKey2 = "+{F9}"

1         On Error GoTo ErrHandler

2         If SwitchOn Then
3             If ActiveSheet Is shCreditUsage Or _
                  ActiveSheet Is shTable Or _
                  ActiveSheet Is shScenarioDefinition Or _
                  ActiveSheet Is shScenarioResults Or _
                  ActiveSheet Is shBarChart Or _
                  ActiveSheet Is shHistoricalData Or _
                  ActiveSheet Is shScenarioCompare Or _
                  ActiveSheet Is shTradesViewer Or _
                  ActiveSheet Is shExportToTMS Then
4                 Application.OnKey TheKey, "F8Response"
5             Else
6                 Application.OnKey TheKey
7             End If
8             If ActiveSheet Is shCreditUsage Or _
                  ActiveSheet Is shScenarioDefinition Then
9                 Application.OnKey TheKey2, "ShiftF9Response"
10            Else
11                Application.OnKey TheKey2
12            End If
13        Else
14            If Not ActiveWindow Is Nothing Then 'This tests avoids errors when the the active thingy is not a Window but a ProtectedView window
15                Application.OnKey TheKey
16                Application.OnKey TheKey2
17            End If
18        End If

19        Exit Sub
ErrHandler:
20        Throw "#SetKeys (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub testAMB()
1         On Error GoTo ErrHandler
2         AlignMenuButtons True
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#testAMB (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AlignMenuButtons
' Author    : Philip Swannell
' Date      : 17-Jan-2017
' Purpose   : Align the "Menu" buttons on all sheets.
' -----------------------------------------------------------------------------------------------------------------------
Sub AlignMenuButtons(Optional Reset As Boolean)
          Const T As Double = 6
          Const l As Double = 183
          Const W As Double = 90
          Const H As Double = 22.5
          Dim hNudge As Double
          Dim i As Long
          Dim SPH As clsSheetProtectionHandler
          Dim vNudge As Double
          Dim ws As Worksheet
          Const ButtonCaption = "Menu... (F8)"
          Const ButtonCaption2 = "Export... (F8)"

1         On Error GoTo ErrHandler
2         For i = 1 To 9
3             Set ws = Choose(i, shBarChart, shCreditUsage, shTable, _
                  shTradesViewer, shScenarioDefinition, shScenarioResults, shScenarioCompare, shHistoricalData, shExportToTMS)
4             If ws Is shHistoricalData Then
5                 hNudge = -l + 5
6                 vNudge = -T + 29.25
7             ElseIf ws Is shExportToTMS Then
8                 hNudge = -l + shExportToTMS.Cells(1, 1).Width
9             Else
10                hNudge = 0
11                vNudge = 0
12            End If
              Dim b As Button
13            If IsInCollection(ws.Buttons, "butMenu") Then
14                Set b = ws.Buttons("butMenu")
15                If Reset Or b.Top <> T + vNudge Or b.Left <> l + hNudge Or b.Width <> W Or b.Height <> H Then
16                    Set SPH = CreateSheetProtectionHandler(ws)
17                    b.Top = T + vNudge: b.Left = l + hNudge: b.Height = H: b.Width = W
18                    b.Placement = xlFreeFloating
19                    If ws Is shExportToTMS Then
20                        b.Characters.Text = ButtonCaption2
21                    Else
22                        b.Characters.Text = ButtonCaption
23                    End If
24                    With b.Characters.Font
25                        .Name = "Calibri"
26                        .ColorIndex = 48
27                        .Size = 10
28                    End With
29                    With b.Characters(Start:=1, Length:=Len(b.Caption) - 4).Font
30                        .Name = "Calibri"
31                        .ColorIndex = 1
32                        .Size = 14
33                    End With
34                End If
35            End If
36        Next i
37        Exit Sub
ErrHandler:
38        Throw "#AlignMenuButtons (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowVersionInfo
' Author     : Philip Swannell
' Date       : 07-Apr-2022
' Purpose    : Show information about the version of the software in use
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowVersionInfo()

          Dim Component As String
          Dim DataToShow As Variant
          Dim i As Long
          Dim LastModified
          Dim Location As String
          Dim Prompt As String
          Dim Version As Variant
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         DataToShow = sArrayRange("Component", "Version", "LastModified", "Location")

3         For i = 1 To 6
4             Select Case i
                  Case 1
5                     Component = ThisWorkbook.Name
6                     Version = RangeFromSheet(shAudit, "Headers").Cells(2, 1)
7                     Location = ThisWorkbook.Path
8                     LastModified = Format(sFileInfo(ThisWorkbook.FullName, "M"), "dd-mmm-yyyy hh:mm")
9                 Case 2
10                    Set wb = Nothing
11                    On Error Resume Next
12                    Set wb = OpenMarketWorkbook(True, False)
13                    On Error GoTo ErrHandler
14                    Component = "Market Data Workbook"
15                    If wb Is Nothing Then
16                        Version = "#Error!"
17                        Location = sJoinPath(ThisWorkbook.Path, RangeFromSheet(shConfig, "MarketDataWorkbook").Value)
18                    Else
19                        Version = RangeFromSheet(wb.Worksheets("Audit"), "Headers").Cells(2, 1)
20                        Location = wb.FullName
21                    End If
22                    LastModified = Format(sFileInfo(Location, "M"), "dd-mmm-yyyy hh:mm")
23                Case 3
24                    Set wb = Nothing
25                    On Error Resume Next
26                    Set wb = OpenLinesWorkbook(True, False)
27                    On Error GoTo ErrHandler
28                    If wb Is Nothing Then
29                        Version = "#Error!"
30                        Location = sJoinPath(ThisWorkbook.Path, RangeFromSheet(shConfig, "LinesWorkbook").Value)
31                    Else
32                        Component = "Lines Workbook"
33                        Version = RangeFromSheet(wb.Worksheets("Audit"), "Headers").Cells(2, 1)
34                        Location = wb.FullName
35                    End If
36                    LastModified = Format(sFileInfo(Location, "M"), "dd-mmm-yyyy hh:mm")
37                Case 4
38                    Set wb = Application.Workbooks("SolumAddin.xlam")
39                    Component = wb.Name
40                    Version = RangeFromSheet(wb.Worksheets("Audit"), "Headers").Cells(2, 1)
41                    Location = wb.Path
42                    LastModified = Format(sFileInfo(wb.FullName, "M"), "dd-mmm-yyyy hh:mm")
43                Case 5
44                    Set wb = Application.Workbooks("SolumSCRiPTUtils.xlam")
45                    Component = wb.Name
46                    Version = RangeFromSheet(wb.Worksheets("Audit"), "Headers").Cells(2, 1)
47                    Location = wb.Path
48                    LastModified = Format(sFileInfo(wb.FullName, "M"), "dd-mmm-yyyy hh:mm")
49                Case 6
50                    Component = "Julia System Image"
51                    Location = IIf(UseLinux(), gSysImageXVALinux, gSysImageXVAWindows)
52                    Version = ""
53                    LastModified = Format(sFileInfo(Location, "M"), "dd-mmm-yyyy hh:mm")
54            End Select

55            DataToShow = sArrayStack(DataToShow, sArrayRange(Component, Version, LastModified, Location))
56        Next i

57        Prompt = "Cayley2022 version information." & vbLf & vbLf & _
              sJustifyArrayOfStrings(DataToShow, , , vbTab, , , True) & vbLf & vbLf & _
              "Julia Status:" & vbLf & _
              JuliaStatus()

58        If MsgBoxPlus(Prompt, vbInformation + vbOKCancel + vbDefaultButton2, "Cayley2022", "Copy to clipboard", "OK", , , 600) = vbOK Then
59            CopyStringToClipboard Prompt
60        End If

61        Exit Sub
ErrHandler:
62        Throw "#ShowVersionInfo (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Function JuliaStatus()

1         On Error GoTo ErrHandler
2         JuliaLaunchForCayley
3         JuliaStatus = JuliaEvalVBA("using Pkg;io=IOBuffer();Pkg.status(io=io);String(take!(io))")

4         Exit Function
ErrHandler:
5         JuliaStatus = "#JuliaStatus (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function CheckAddinVersion(AddinName As String, Optional MinimumVersionRequired As Long)

          Dim ErrorDescription As String
          Dim HaveVersion As Long
1         On Error GoTo ErrHandler

2         HaveVersion = RangeFromSheet(Application.Workbooks(AddinName).Worksheets("Audit"), "Headers").Cells(2, 1).Value
3         If HaveVersion < MinimumVersionRequired Then
4             ErrorDescription = "This is version " & Format(RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value, "###,###") & " of " & ThisWorkbook.Name & ". It requires version " & _
                  Format(MinimumVersionRequired, "###,###") & " of the Excel addin " & AddinName & ". However, you are using an older version (" & Format(HaveVersion, "###,###") & ") installed at " & _
                  Application.Workbooks(AddinName).FullName & vbLf & vbLf & _
                  "You probably need to re-install the Cayley software. Visit " & gGitHubRepo & " using GitHub account " & gGitHubAccount & ". You will need a password to access that site."
          
5             Throw ErrorDescription

6         End If

7         Exit Function
ErrHandler:
8         Throw "#CheckAddinVersion (line " & CStr(Erl) & "): " & Err.Description & "!", True
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SafeAppActivate
' Author     : Philip Swannell
' Date       : 23-Apr-2022
' Purpose    : AppActivate Application.Caption sometimes fails, but we don't want an error to be thrown, so wrap with
'              On Error Resume Next
' -----------------------------------------------------------------------------------------------------------------------
Sub SafeAppActivate(SheetToActivate As Worksheet)
          Dim EN As Long
1         On Error Resume Next
          SheetToActivate.Activate
2         AppActivate Application.Caption
End Sub

