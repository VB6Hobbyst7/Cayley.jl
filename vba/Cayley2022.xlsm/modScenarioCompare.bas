Attribute VB_Name = "modScenarioCompare"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modScenarioCompare
' Author    : Philip Swannell
' Date      : 28-Feb-2017
' Purpose   : Code relating to the CompareScenarios sheet. The code of this module would
'             benefit from restructuring, but time does not allow. For example, there is
'             repeated code for constructing a chart to display a scenario in methods
'             RefreshScenarioResultsSheet and RefreshScenarioCompareSheet, and the latter
'             method is too long and ought to be refactored...
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuScenarioCompare
' Author    : Philip Swannell
' Date      : 28-Feb-2017
' Purpose   : Attached to menu in ScenarioCompare sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuScenarioCompare()

          Const chCompare = "&Compare two Scenarios..."
          Const FidCompare = 5898
          Dim Res As String

1         On Error GoTo ErrHandler
          
2         RunThisAtTopOfCallStack
          
3         Res = ShowCommandBarPopup(chCompare, FidCompare, , , ChooseAnchorObject())
4         Select Case Res
              Case "#Cancel!"
5             Case Unembellish(chCompare)
6                 OpenTwoScenariosWithPrompt
7             Case Else
8                 Throw "Unrecognised choice"
9         End Select

10        shScenarioCompare.Protect , True, True

11        Exit Sub
ErrHandler:
12        SomethingWentWrong "#MenuScenarioCompare (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenTwoScenariosWithPrompt
' Author    : Philip Swannell
' Date      : 28-Feb-2017
' Purpose   : Wrapper to OpenTwoScenarioFiles with friendly message box...
' -----------------------------------------------------------------------------------------------------------------------
Sub OpenTwoScenariosWithPrompt()
1         On Error GoTo ErrHandler

          Dim Prompt
          Dim Res As VbMsgBoxResult
          Static CheckBoxValue As Boolean
          Const Title = "Compare Two Scenarios"

2         If Not CheckBoxValue Then
3             Prompt = "This method allows you to compare the results of two scenarios and to see the difference " & _
                  "between hedge capacities during their life. This is useful for looking at the impact of various" & _
                  " trading strategies or speed grids. " & vbLf & vbLf & _
                  "First select the scenario results file for Scenario 1 (the ""Base"" scenario) then select the " & _
                  "file for Scenario 2. The sheet displays a comparison between the two scenario definitions " & _
                  "(differences highlighted in yellow), the two scenario graphs and a graph showing the " & _
                  "differences between the hedge capacities." & vbLf & vbLf & _
                  "Proceed?"
4             Res = MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, , , , , , _
                  "Don't show this message again", CheckBoxValue)

5             If Res <> vbOK Then
6                 CheckBoxValue = False
7                 Exit Sub
8             End If
9         End If

10        OpenTwoScenarioFiles "", ""

11        Exit Sub
ErrHandler:
12        SomethingWentWrong "#OpenTwoScenariosWithPrompt (line " & CStr(Erl) & "): " & Err.Description & "!"

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenTwoScenarioFiles
' Author    : Philip Swannell
' Date      : 27-Feb-2017
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub OpenTwoScenarioFiles(FileName1 As String, FileName2 As String)
          Dim D1 As Dictionary
          Dim D2 As Dictionary

          Dim CopyOfErr As String
          Dim FileFilter As String
          Dim oldBlockChange As Boolean
          Dim Prompt As String
          Dim RegKey As String
          Dim Res As Variant
          Dim ShowSwitchOption As Boolean
          Dim SUH As clsScreenUpdateHandler
          Dim Title As String
          Dim Tmp As String

1         On Error GoTo ErrHandler

2         oldBlockChange = gBlockChangeEvent
3         gBlockChangeEvent = True

4         RegKey = gRegKey_Res
5         FileFilter = "ScenarioResultsFile (*.srf),*.srf"

6         If FileName1 = "" Then
TryAgain:
7             Title = "Select First or Both Scenario Results File(s)"
8             Res = GetOpenFilenameWrap(RegKey, FileFilter, 1, Title, , True, False)
9             If VarType(Res) = vbBoolean Then
10                GoTo EarlyExit
11            ElseIf sNCols(Res) = 1 Then
12                FileName1 = Res(1, 1)
13            ElseIf sNCols(Res) = 2 Then
14                ShowSwitchOption = True
15                FileName1 = Res(1, 1)
16                FileName2 = Res(1, 2)
17            Else
18                GoTo TryAgain
19            End If
20        End If

21        If ShowSwitchOption Then
22            Title = "Compare Scenarios"
23            Prompt = "Is this the correct order for the scenarios?" + vbLf + _
                  "Scenario1 =       " & sSplitPath(FileName1) + vbLf + _
                  "Scenario2 =       " & sSplitPath(FileName2)
24            Select Case MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, Title, _
                  "Yes, Compare them", "Flip 1 and 2, then Compare", "Quit")
                  Case vbNo
25                    Tmp = FileName1
26                    FileName1 = FileName2
27                    FileName2 = Tmp
28                Case vbCancel
29                    GoTo EarlyExit
30            End Select
31        End If

32        If FileName2 = "" Then
33            Title = "Select Second Scenario Results File"
34            FileName2 = GetOpenFilenameWrap(RegKey, FileFilter, 1, Title, , False, False)
35            If FileName2 = "False" Then GoTo EarlyExit
36        End If

37        Set SUH = CreateScreenUpdateHandler()

          'Gets added back to MRU at the end of the method if all is well with the files
38        RemoveFileFromMRU RegKey, FileName1
39        RemoveFileFromMRU RegKey, FileName2

40        Set D1 = SRFFileToDictionary(FileName1)
41        D1.Add "FileName", FileName1

42        Set D2 = SRFFileToDictionary(FileName2)
43        D2.Add "FileName", FileName2

44        gBlockChangeEvent = oldBlockChange

45        RefreshScenarioCompareSheet D1, D2
          'Add file to Most Recently Used list only if the opening was successful
46        AddFileToMRU RegKey, FileName1
47        AddFileToMRU RegKey, FileName2

EarlyExit:
48        gBlockChangeEvent = oldBlockChange

49        Exit Sub
ErrHandler:
50        CopyOfErr = "#OpenTwoScenarioFiles (line " & CStr(Erl) & "): " & Err.Description & "!"
51        gBlockChangeEvent = oldBlockChange
52        Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SRFFileToDictionary
' Author    : Philip Swannell
' Date      : 27-Feb-2017
' Purpose   : Opens an scenario results file, populates a dictionary and closes the file again, does not
'             feed the data from the file into the workbook (so differs from method OpenScenarioFile)
' -----------------------------------------------------------------------------------------------------------------------
Function SRFFileToDictionary(FileName As String) As Dictionary
          Dim CopyOfErr As String
          Dim CurrentVersion As Long
          Dim D As New Dictionary
          Dim i As Long
          Dim IsDefinitionInFile As Variant
          Dim Map As Variant
          Dim OverwriteWith As Variant
          Dim RangeName As String
          Dim SheetName As String
          Dim SourceSheet As Worksheet
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         CurrentVersion = RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value

          'Close it if it's already open  (because of an earlier error?)
3         If IsInCollection(Application.Workbooks, sSplitPath(FileName)) Then
4             Application.Workbooks(sSplitPath(FileName)).Close False
5         End If

6         If Not sFileExists(FileName) Then
7             Throw "Sorry, we couldn't find '" & FileName & "'" & vbLf & _
                  "Is it possible it was moved, renamed or deleted?", True
8         End If

9         Set wb = Application.Workbooks.Open(FileName, , , , "Foo")
10        Set SourceSheet = wb.Worksheets(1)
11        If Not IsInCollection(SourceSheet.Names, "Map") Then Throw "Unexpected error, cannot find 'Map' in file. So the file is not a valid Scenario Results file"
12        Set Map = RangeFromSheet(SourceSheet, "Map")

          Dim SavedByVersion

13        SavedByVersion = sVLookup("SavedByVersion", Map.Value)
14        If Not IsNumber(SavedByVersion) Then
15            Throw FileName & " was saved by an old version of the Cayley workbook that used a file format incompatible with the current version of the Cayley workbook. Sorry, but you cannot use it.", True
16        ElseIf SavedByVersion < gOldestSupportedScenarioVersion Then
17            Throw FileName & " was saved by version " & CStr(SavedByVersion) & " of the Cayley workbook. Unfortunately, this version of the Cayley workbook (" & CStr(CurrentVersion) & ") is not compatible with files that old and you cannot open them.", True
18        End If

19        IsDefinitionInFile = sVLookup("IsDefinition", Map.Value)
20        If Not VarType(IsDefinitionInFile) = vbBoolean Then
21            Throw FileName & " was saved by an old version of the Cayley workbook that used a file format incompatible with the current version of the Cayley workbook. Sorry"
22        ElseIf Not sEquals(IsDefinitionInFile, False) Then
23            Throw "That file is a Scenario Definition file, but you need to open a Scenario Results file"
24        End If

25        For i = 1 To Map.Rows.Count - 2
26            SheetName = Map.Cells(i, 1).Value
27            RangeName = Map.Cells(i, 2).Value
28            MorphRangeAndSheetNames SheetName, OverwriteWith
29            If SheetName = shScenarioResults.Name Then
30                D.Add RangeName, SourceSheet.Range(Map(i, 3)).Value
31            End If
32        Next i
33        D.Add "SavedByVersion", SavedByVersion

          'For backward compatibility with Cayley 2017 vintage
34        If Not D.Exists("HedgeHorizon") Then
35            D.Add "HedgeHorizon", (sNCols(D("ScenarioResultsHeaders")) - 1) / 2
36        End If

37        Set SRFFileToDictionary = D
38        wb.Close False
39        Exit Function
ErrHandler:
40        CopyOfErr = "#SRFFileToDictionary (line " & CStr(Erl) & "): " & Err.Description & "!"
41        If Not wb Is Nothing Then
42            wb.Close False
43        End If
44        Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshScenarioCompareSheet
' Author    : Philip Swannell
' Date      : 28-Feb-2017
' Purpose   : Refreshes the ScenarioCompare worksheet. Left three columns are updated,
'             to the right of that the worksheet is reconstructed from scratch. The two
'             dictionary arguments will have been created by calls to SRFFileToDictionary.
'             This method would benefit from refactoring (and sharing sub-routines with
'             RefreshScenarioResultsSheet).
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshScenarioCompareSheet(D1 As Dictionary, D2 As Dictionary)

          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

          'Chart Legends
          Const L_HedgeCapacity = "Hedge Capacity ($ bln, right axis)"
          Const L_Spot = "EURUSD (left axis)"
          Const L_Vol = "3Y EURUSD Vol  (x 10, left axis)"
          Const L_HedgeCompetion = "Hedge completion ratio (%, right axis)"
          Const L_LineEx = "Line exhaustion ($ bln, right axis)"
          Const L_Time = "Time (months)"
          Const StartAddresss = "AA3"        'We paste in the data for plotting starting at this address
          Const TopLeftAddress1 = "E30"        'top left of chart goes here
          Const TopLeftAddress2 = "E57"        'top left of chart goes here
          Const TopLeftAddress3 = "E3"        'top left of chart goes here
          Const KPIAddress1 = "R30"        'KPIs goes here
          Const KPIAddress2 = "R57"        'KPIs goes here
          Const KPIAddressDiffs = "R3"

          Dim c1 As Range
          Dim c2 As Range
          Dim i As Long
          Dim NR1 As Long
          Dim NR2 As Long
          Dim ThisName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set ws = shScenarioCompare
3         Set SPH = CreateSheetProtectionHandler(ws)
4         Set SUH = CreateScreenUpdateHandler()

          Dim NamesToPopulate As Variant

5         NamesToPopulate = sTokeniseString("FileName,ShocksDerivedFrom,HistoryStart,HistoryEnd,BaseSpot,BaseVol," & _
              "ForwardsRatio,PutRatio,CallRatio,PutStrikeOffset,CallStrikeOffset,StrategySwitchPoints,HedgeHorizon," & _
              "AllocationByYear,UseSpeedGrid,SpeedGridWidth,HighFxSpeed,LowFxSpeed,VaryGridWidth,SpeedGridBaseVol," & _
              "AnnualReplenishment,ModelType,NumMCPaths,NumObservations,FilterBy2,Filter2Value,IncludeAssetClasses," & _
              "CurrenciesToInclude,TradesScaleFactor,LinesScaleFactor,TimeStart,TimeEnd,ComputerName," & _
              "ScenarioDescription,SavedByVersion")

          'Populate the left three columns of the sheet
6         For i = 1 To sNRows(NamesToPopulate)
7             ThisName = NamesToPopulate(i, 1)
8             Set c1 = RangeFromSheet(ws, ThisName & "1")
9             Set c2 = RangeFromSheet(ws, ThisName & "2")

10            c1.Value = sArrayExcelString(DictGet(D1, CStr(ThisName)))
11            c2.Value = sArrayExcelString(DictGet(D2, CStr(ThisName)))

12            If c1.Value = c2.Value Then
13                c1.Interior.ColorIndex = xlColorIndexAutomatic
14                c2.Interior.ColorIndex = xlColorIndexAutomatic
15            Else
16                Select Case ThisName
                      Case "TimeStart", "TimeEnd", "ScenarioDescription", "FileName"    'differences to be expected
17                        c1.Interior.ColorIndex = xlColorIndexAutomatic
18                        c2.Interior.ColorIndex = xlColorIndexAutomatic
19                    Case Else
20                        c1.Interior.Color = 6750207
21                        c2.Interior.Color = 6750207
22                End Select
23            End If
24        Next i

          Dim HedgeHorizon1
          Dim HedgeHorizon2
          Dim ScenarioDefinition1
          Dim ScenarioDefinition2
          Dim ScenarioDefinitionHeaders1
          Dim ScenarioDefinitionHeaders2
          Dim ScenarioResults1
          Dim ScenarioResults2
          Dim ScenarioResultsHeaders1
          Dim ScenarioResultsHeaders2

25        ScenarioDefinitionHeaders1 = DictGet(D1, "ScenarioDefinitionHeaders")
26        ScenarioDefinition1 = DictGet(D1, "ScenarioDefinition")
27        ScenarioResultsHeaders1 = DictGet(D1, "ScenarioResultsHeaders")
28        ScenarioResults1 = DictGet(D1, "ScenarioResults")
29        HedgeHorizon1 = DictGet(D1, "HedgeHorizon")

30        ScenarioDefinitionHeaders2 = DictGet(D2, "ScenarioDefinitionHeaders")
31        ScenarioDefinition2 = DictGet(D2, "ScenarioDefinition")
32        ScenarioResultsHeaders2 = DictGet(D2, "ScenarioResultsHeaders")
33        ScenarioResults2 = DictGet(D2, "ScenarioResults")
34        HedgeHorizon2 = DictGet(D2, "HedgeHorizon")

35        Force2DArrayRMulti ScenarioDefinitionHeaders1, ScenarioDefinition1, ScenarioResultsHeaders1, ScenarioResults1
36        Force2DArrayRMulti ScenarioDefinitionHeaders2, ScenarioDefinition2, ScenarioResultsHeaders2, ScenarioResults2

37        If sNRows(ScenarioDefinitionHeaders1) = 1 Then ScenarioDefinitionHeaders1 = sArrayTranspose(ScenarioDefinitionHeaders1)
38        If sNRows(ScenarioResultsHeaders1) = 1 Then ScenarioResultsHeaders1 = sArrayTranspose(ScenarioResultsHeaders1)
39        If sNRows(ScenarioDefinitionHeaders2) = 1 Then ScenarioDefinitionHeaders2 = sArrayTranspose(ScenarioDefinitionHeaders2)
40        If sNRows(ScenarioResultsHeaders2) = 1 Then ScenarioResultsHeaders2 = sArrayTranspose(ScenarioResultsHeaders2)

41        If Not sArraysIdentical(ScenarioDefinitionHeaders1, ScenarioDefinitionHeaders2) Then
42            Throw "ScenarioDefinitionHeaders not consistent between the two files"
43        End If
44        If Not sArraysIdentical(ScenarioResultsHeaders1, ScenarioResultsHeaders2) Then
45            Throw "ScenarioResultsHeaders not consistent between the two files"
46        End If

47        NR1 = sNRows(ScenarioDefinition1)
48        NR2 = sNRows(ScenarioDefinition2)

          'Very basic error checking...
49        If sNRows(ScenarioResults1) <> NR1 Then Throw "ScenarioDefinition1 and ScenarioResults1 must have the same number of rows " & _
              "but ScenarioDefinition has " & CStr(sNRows(ScenarioDefinition1)) & " rows and ScenarioResults has " & CStr(sNRows(ScenarioResults1)) & " rows."

50        If sNRows(ScenarioResults2) <> NR2 Then Throw "ScenarioDefinition2 and ScenarioResults2 must have the same number of rows " & _
              "but ScenarioDefinition has " & CStr(sNRows(ScenarioDefinition2)) & " rows and ScenarioResults has " & CStr(sNRows(ScenarioResults2)) & " rows."

          'Read headers to get column numbers, can assume
          Dim cnd_AvEDSTraded
          Dim cnd_FxShock
          Dim cnd_FxVolShock
          Dim cnd_Months
          Dim cnd_ReplenishmentAmount
          Dim cnr_1YC
          Dim cnr_1YT
          Dim cnr_HHYC1
          Dim cnr_HHYC2
          Dim cnr_HHYT1
          Dim cnr_HHYT2

51        cnr_1YT = ThrowIfError(sMatch("1Y Traded", ScenarioResultsHeaders1))
52        cnr_HHYT1 = ThrowIfError(sMatch(CStr(HedgeHorizon1) & "Y Traded", ScenarioResultsHeaders1))
53        cnr_HHYT2 = ThrowIfError(sMatch(CStr(HedgeHorizon2) & "Y Traded", ScenarioResultsHeaders2))

54        cnr_1YC = ThrowIfError(sMatch("1Y capacity", ScenarioResultsHeaders1))
55        cnr_HHYC1 = ThrowIfError(sMatch(CStr(HedgeHorizon1) & "Y capacity", ScenarioResultsHeaders1))
56        cnr_HHYC2 = ThrowIfError(sMatch(CStr(HedgeHorizon2) & "Y capacity", ScenarioResultsHeaders2))

57        cnd_Months = ThrowIfError(sMatch("Months", ScenarioDefinitionHeaders1))
58        cnd_FxShock = ThrowIfError(sMatch("FxShock", ScenarioDefinitionHeaders1))
59        cnd_FxVolShock = ThrowIfError(sMatch("FxVolShock", ScenarioDefinitionHeaders1))
60        cnd_ReplenishmentAmount = ThrowIfError(sMatch("ReplenishmentAmount", ScenarioDefinitionHeaders1))
61        cnd_AvEDSTraded = sMatch("AvEDSTraded", ScenarioDefinitionHeaders1)

          'Calculate what we want to plot in the two charts
          Dim BaseSpot1
          Dim BaseSpot2
          Dim BaseVol1
          Dim BaseVol2
          Dim ChooseVector1
          Dim ChooseVector2
          Dim DatesForPlotting1
          Dim DatesForPlotting2
          Dim HedgeCapacity1
          Dim HedgeCapacity2
          Dim HedgeCompletionRatio1
          Dim HedgeCompletionRatio2
          Dim HistoryDates1
          Dim HistoryDates2
          Dim HistoryEnd1
          Dim HistoryEnd2
          Dim HistorySpot1
          Dim HistorySpot2
          Dim HistoryStart1
          Dim HistoryStart2
          Dim HistoryVol1
          Dim HistoryVol2
          Dim isHighRes1 As Boolean
          Dim isHighRes2 As Boolean
          Dim LineExhaustionLevel1
          Dim LineExhaustionLevel2
          Dim MaxMonth1
          Dim MaxMonth2
          Dim ShocksDerivedFrom1 As String
          Dim ShocksDerivedFrom2 As String
          Dim SpotForPlotting1
          Dim SpotForPlotting2
          Dim TargetTraded1toHH1
          Dim TargetTraded1toHH2
          Dim TimeForPlotting1
          Dim TimeForPlotting2
          Dim TotalTraded1toHH1
          Dim TotalTraded1toHH2
          Dim VolForPlotting1
          Dim VolForPlotting2

62        HedgeCapacity1 = sRowSum(sSubArray(ScenarioResults1, 1, cnr_1YT, , cnr_HHYC1 - cnr_1YT + 1))
63        HedgeCapacity1 = sArrayDivide(HedgeCapacity1, 1000000000#)
64        HedgeCapacity2 = sRowSum(sSubArray(ScenarioResults2, 1, cnr_1YT, , cnr_HHYC2 - cnr_1YT + 1))
65        HedgeCapacity2 = sArrayDivide(HedgeCapacity2, 1000000000#)

66        TotalTraded1toHH1 = sRowSum(sSubArray(ScenarioResults1, 1, cnr_1YT, , cnr_HHYT1 - cnr_1YT + 1))
67        TotalTraded1toHH2 = sRowSum(sSubArray(ScenarioResults2, 1, cnr_1YT, , cnr_HHYT2 - cnr_1YT + 1))
68        TargetTraded1toHH1 = sSubArray(ScenarioDefinition1, 1, cnd_ReplenishmentAmount, , 1)
69        TargetTraded1toHH2 = sSubArray(ScenarioDefinition2, 1, cnd_ReplenishmentAmount, , 1)

70        HedgeCompletionRatio1 = sArrayMultiply(sArrayDivide(sPartialSum(TotalTraded1toHH1), sPartialSum(TargetTraded1toHH1)), 100)
71        HedgeCompletionRatio2 = sArrayMultiply(sArrayDivide(sPartialSum(TotalTraded1toHH2), sPartialSum(TargetTraded1toHH2)), 100)

72        LineExhaustionLevel1 = sReshape(10, NR1, 1)
73        LineExhaustionLevel2 = sReshape(10, NR2, 1)

74        ShocksDerivedFrom1 = DictGet(D1, "ShocksDerivedFrom")
75        ShocksDerivedFrom2 = DictGet(D2, "ShocksDerivedFrom")
76        HistoryStart1 = DictGet(D1, "HistoryStart")
77        HistoryStart2 = DictGet(D2, "HistoryStart")
78        HistoryEnd1 = DictGet(D1, "HistoryEnd")
79        HistoryEnd2 = DictGet(D2, "HistoryEnd")
80        BaseSpot1 = DictGet(D1, "BaseSpot")
81        BaseSpot2 = DictGet(D2, "BaseSpot")
82        BaseVol1 = DictGet(D1, "BaseVol")
83        BaseVol2 = DictGet(D2, "BaseVol")
84        isHighRes1 = LCase(ShocksDerivedFrom1) = "history"
85        isHighRes2 = LCase(ShocksDerivedFrom2) = "history"
86        MaxMonth1 = ScenarioDefinition1(sNRows(ScenarioDefinition1), cnd_Months)
87        MaxMonth2 = ScenarioDefinition2(sNRows(ScenarioDefinition2), cnd_Months)

88        If isHighRes1 Then
89            With RangeFromSheet(shHistoricalData, "TheDates")
90                HistoryDates1 = .Value
91                HistorySpot1 = .offset(, 1).Value
92                HistoryVol1 = .offset(, 2).Value
93            End With
94            ChooseVector1 = sArrayAnd(sArrayGreaterThanOrEqual(HistoryDates1, HistoryStart1), _
                  sArrayLessThanOrEqual(HistoryDates1, HistoryEnd1))
95            DatesForPlotting1 = sMChoose(HistoryDates1, ChooseVector1)
96            SpotForPlotting1 = sMChoose(HistorySpot1, ChooseVector1)
97            SpotForPlotting1 = sArrayMultiply(SpotForPlotting1, BaseSpot1 / SpotForPlotting1(1, 1))
98            VolForPlotting1 = sMChoose(HistoryVol1, ChooseVector1)
99            TimeForPlotting1 = sArraySubtract(DatesForPlotting1, HistoryStart1)
100           TimeForPlotting1 = sArrayMultiply(TimeForPlotting1, MaxMonth1 / (HistoryEnd1 - HistoryStart1))
101       Else
102           TimeForPlotting1 = sSubArray(ScenarioDefinition1, 1, cnd_Months, , 1)
103           SpotForPlotting1 = sArrayMultiply(sSubArray(ScenarioDefinition1, 1, cnd_FxShock, , 1), BaseSpot1)
104           VolForPlotting1 = sArrayMultiply(sSubArray(ScenarioDefinition1, 1, cnd_FxVolShock, , 1), BaseVol1)
105       End If
106       VolForPlotting1 = sArrayMultiply(VolForPlotting1, 10)        'To scale to fit on the chart

107       If isHighRes2 Then
108           With RangeFromSheet(shHistoricalData, "TheDates")
109               HistoryDates2 = .Value
110               HistorySpot2 = .offset(, 1).Value
111               HistoryVol2 = .offset(, 2).Value
112           End With
113           ChooseVector2 = sArrayAnd(sArrayGreaterThanOrEqual(HistoryDates2, HistoryStart2), _
                  sArrayLessThanOrEqual(HistoryDates2, HistoryEnd2))
114           DatesForPlotting2 = sMChoose(HistoryDates2, ChooseVector2)
115           SpotForPlotting2 = sMChoose(HistorySpot2, ChooseVector2)
116           SpotForPlotting2 = sArrayMultiply(SpotForPlotting2, BaseSpot2 / SpotForPlotting2(1, 1))
117           VolForPlotting2 = sMChoose(HistoryVol2, ChooseVector2)
118           TimeForPlotting2 = sArraySubtract(DatesForPlotting2, HistoryStart2)
119           TimeForPlotting2 = sArrayMultiply(TimeForPlotting2, MaxMonth2 / (HistoryEnd2 - HistoryStart2))
120       Else
121           TimeForPlotting2 = sSubArray(ScenarioDefinition2, 1, cnd_Months, , 1)
122           SpotForPlotting2 = sArrayMultiply(sSubArray(ScenarioDefinition2, 1, cnd_FxShock, , 1), BaseSpot2)
123           VolForPlotting2 = sArrayMultiply(sSubArray(ScenarioDefinition2, 1, cnd_FxVolShock, , 1), BaseVol2)
124       End If
125       VolForPlotting2 = sArrayMultiply(VolForPlotting2, 10)        'To scale to fit on the chart

          'Clear out the sheet, apart from the left two columns
          Dim o As Object
          Dim Res
126       For Each o In ws.ChartObjects
127           o.Delete
128       Next
          Dim N As Name
          Dim R As Range
129       For Each N In ws.Names
130           Set R = Nothing
131           On Error Resume Next
132           Set R = N.RefersToRange
133           On Error GoTo ErrHandler
134           If R Is Nothing Then
135               N.Delete
136           ElseIf R.Column > 3 Then
137               If R.Parent Is ws Then
138                   N.Delete
139               End If
140           End If
141       Next N
142       ws.Cells(1, 4).Resize(1, 1000).EntireColumn.Delete
143       Res = ws.UsedRange.Rows.Count        'resets used range

          'Paste in the data for plotting...
          Dim StartCell As Range
          Dim TargetRange As Range

144       For i = 1 To 2
145           If i = 1 Then
146               Set StartCell = ws.Range(StartAddresss)
147           Else
148               With TargetRange
149                   Set StartCell = ws.Cells(ws.Range(StartAddresss).Row, .Column + .Columns.Count + 1)
150               End With
151           End If

              Dim HedgeCapacity
              Dim HedgeCompletionRatio
              Dim LineExhaustionLevel
              Dim ScenarioDefinition
              Dim ScenarioDefinitionHeaders
              Dim ScenarioResults
              Dim ScenarioResultsHeaders
              Dim SpotForPlotting
              Dim TimeForPlotting
              Dim VolForPlotting

152           ScenarioDefinition = Choose(i, ScenarioDefinition1, ScenarioDefinition2)
153           ScenarioDefinitionHeaders = Choose(i, ScenarioDefinitionHeaders1, ScenarioDefinitionHeaders2)
154           ScenarioResults = Choose(i, ScenarioResults1, ScenarioResults2)
155           ScenarioResultsHeaders = Choose(i, ScenarioResultsHeaders1, ScenarioResultsHeaders2)
156           HedgeCapacity = Choose(i, HedgeCapacity1, HedgeCapacity2)
157           HedgeCompletionRatio = Choose(i, HedgeCompletionRatio1, HedgeCompletionRatio2)
158           LineExhaustionLevel = Choose(i, LineExhaustionLevel1, LineExhaustionLevel2)
159           TimeForPlotting = Choose(i, TimeForPlotting1, TimeForPlotting2)
160           SpotForPlotting = Choose(i, SpotForPlotting1, SpotForPlotting2)
161           VolForPlotting = Choose(i, VolForPlotting1, VolForPlotting2)

162           Set TargetRange = StartCell.Resize(sNRows(ScenarioDefinition) + 1, sNCols(ScenarioDefinition))
163           With TargetRange
164               .Value = sArrayStack(sArrayTranspose(ScenarioDefinitionHeaders), ScenarioDefinition)
165               .Rows(1).Font.Bold = True
166               AddGreyBorders .offset(0), True
167               ws.Names.Add "ScenarioDefinitionHeaders" & CStr(i), .Rows(1)
168               ws.Names.Add "ScenarioDefinition" & CStr(i), .offset(1).Resize(.Rows.Count - 1)
169               .Columns(cnd_ReplenishmentAmount).NumberFormat = "#,##0;[Red]-#,##0"
170               .Columns.AutoFit
171               .Cells(0, 1).Value = "ScenarioDefinition" & CStr(i)
172               Set TargetRange = .Cells(1, .Columns.Count + 2)
173           End With

              'ScenarioResults...
174           Set TargetRange = TargetRange.Resize(sNRows(ScenarioResults) + 1, sNCols(ScenarioResults))
175           With TargetRange
176               .Value = sArrayStack(sArrayTranspose(ScenarioResultsHeaders), ScenarioResults)
177               .Rows(1).Font.Bold = True
178               AddGreyBorders .offset(0), True
179               ws.Names.Add "ScenarioResultsHeaders" & CStr(i), .Rows(1)
180               ws.Names.Add "ScenarioResults" & CStr(i), .offset(1).Resize(.Rows.Count - 1)
181               .NumberFormat = "#,##0;[Red]-#,##0"
182               .Columns.AutoFit
183               .Cells(0, 1).Value = "ScenarioResults" & CStr(i)
184               .Cells(0, cnr_1YC) = "Capacity after allocation of trades"
185               Set TargetRange = .Cells(1, .Columns.Count + 2)
186           End With

              'HedgeCapacity etc...
187           Set TargetRange = TargetRange.Resize(sNRows(HedgeCapacity) + 1, 1)
188           With TargetRange
189               .Cells(0, 1).Value = "HedgeCapacity" & CStr(i)
190               .Value = sArrayStack(sArrayTranspose(L_HedgeCapacity), HedgeCapacity)
191               AddGreyBorders .offset(0), True
192               ws.Names.Add "HedgeCapacity" & CStr(i), .offset(1).Resize(.Rows.Count - 1)
193               .Columns.AutoFit
194               Set TargetRange = .Cells(1, .Columns.Count + 1)
195           End With

196           Set TargetRange = TargetRange.Resize(sNRows(HedgeCompletionRatio) + 1, 1)
197           With TargetRange
198               .Cells(0, 1).Value = "HedgeCompletionRatio" & CStr(i)
199               .Value = sArrayStack(sArrayTranspose(L_HedgeCompetion), HedgeCompletionRatio)
200               AddGreyBorders .offset(0), True
201               ws.Names.Add "HedgeCompletionRatio" & CStr(i), .offset(1).Resize(.Rows.Count - 1)
202               .Columns.AutoFit
203               Set TargetRange = .Cells(1, .Columns.Count + 1)
204           End With
205           Set TargetRange = TargetRange.Resize(sNRows(LineExhaustionLevel) + 1, 1)
206           With TargetRange
207               .Cells(0, 1).Value = "LineExhaustionLevel" & CStr(i)
208               .Value = sArrayStack(sArrayTranspose(L_LineEx), LineExhaustionLevel)
209               AddGreyBorders .offset(0), True
210               ws.Names.Add "LineExhaustionLevel" & CStr(i), .offset(1).Resize(.Rows.Count - 1)
211               .Columns.AutoFit
212               Set TargetRange = .Cells(2, .Columns.Count + 2)
213           End With
214           Set TargetRange = TargetRange.Resize(sNRows(TimeForPlotting), 1)
215           With TargetRange
216               .Value = TimeForPlotting
217               AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
218               .Cells(0, 1) = L_Time
219               ws.Names.Add "TimeForPlotting" & CStr(i), .offset(0)
220               .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
221               Set TargetRange = .Cells(1, .Columns.Count + 1)
222           End With
223           Set TargetRange = TargetRange.Resize(sNRows(SpotForPlotting), 1)
224           With TargetRange
225               .Value = SpotForPlotting
226               AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
227               .Cells(0, 1) = L_Spot
228               ws.Names.Add "SpotForPlotting" & CStr(i), .offset(0)
229               .NumberFormat = "0.00"
230               .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
231               Set TargetRange = .Cells(1, .Columns.Count + 1)
232           End With
233           Set TargetRange = TargetRange.Resize(sNRows(VolForPlotting), 1)
234           With TargetRange
235               .Value = VolForPlotting
236               AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
237               .Cells(0, 1) = L_Vol
238               ws.Names.Add "VolForPlotting" & CStr(i), .offset(0)
239               .NumberFormat = "0.00"
240               .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
241               Set TargetRange = .Cells(1, .Columns.Count + 2)
242           End With
              '
              'Scenario Performance
              Dim KPIs
              Dim KPIs1
              Dim KPIs2
243           KPIs = sParseArrayString("{""Scenario Performance"",;""Number of stressed periods"",0;""Final Hedge Capacity"",0;""Final completion ratio"",0;""Worst completion ratio"",0}")
244           KPIs(2, 2) = sArrayCount(sArrayLessThanOrEqual(HedgeCapacity, LineExhaustionLevel(1, 1)))
245           KPIs(3, 2) = FirstElement(sTake(HedgeCapacity, -1))
246           KPIs(4, 2) = FirstElement(sTake(HedgeCompletionRatio, -1))
247           KPIs(5, 2) = FirstElement(sColumnMin(HedgeCompletionRatio))
248           If i = 1 Then
249               KPIs1 = KPIs
250           Else
251               KPIs2 = KPIs
252           End If

253           With ws.Range(Choose(i, KPIAddress1, KPIAddress2)).Resize(sNRows(KPIs), sNCols(KPIs))
254               .Value = KPIs
255               .Cells(1, 1).Font.Bold = True
256               .offset(1).Columns.AutoFit
257               .Columns(2).NumberFormat = "0.0"
258               .Cells(2, 2).NumberFormat = "General"
259               AddGreyBorders .offset(0), True
260               ws.Names.Add "KPIs" & CStr(i), .offset(0)
261               .offset(0, 1).Resize(, .Columns.Count - 1).HorizontalAlignment = xlHAlignCenter
262           End With

              Dim MaximumScale As Double
              Dim MinimumScale As Double
263           MinimumScale = sMinOfNums(sArrayStack(SpotForPlotting, VolForPlotting))
264           MinimumScale = Application.WorksheetFunction.Floor(MinimumScale, 0.1)
265           MaximumScale = sMaxOfNums(sArrayStack(SpotForPlotting, VolForPlotting))
266           MaximumScale = Application.WorksheetFunction.Ceiling(MaximumScale, 0.1)

              Dim cht As Chart
              Dim TL As Range

267           Set TL = ws.Range(Choose(i, TopLeftAddress1, TopLeftAddress2))
              'Set up a new chart, using code that's compatible with Excel 2010 i.e. no FullSeriesCollection and no AddChart2
              ' Set cht = ws.Shapes.AddChart2(-1, xlXYScatterLinesNoMarkers, ws.Range(TopLeftAddress).Left, ws.Range(TopLeftAddress).Top, 590, 374.4).Chart
268           Set cht = ws.ChartObjects.Add(Left:=TL.Left, Top:=TL.Top, Width:=590, Height:=374.4).Chart
269           cht.ChartType = xlXYScatterLinesNoMarkers
270           cht.Parent.Placement = xlMove

271           With cht.SeriesCollection.NewSeries
272               .ChartType = xlXYScatterLinesNoMarkers
273               .Name = L_Spot
274               .xValues = "='" & ws.Name & "'!" & RangeFromSheet(ws, "TimeForPlotting" & CStr(i)).Address
275               .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "SpotForPlotting" & CStr(i)).Address
276           End With

277           With cht.SeriesCollection.NewSeries
278               .ChartType = xlXYScatterLinesNoMarkers
279               .Name = L_Vol
280               .xValues = "='" & ws.Name & "'!" & RangeFromSheet(ws, "TimeForPlotting" & CStr(i)).Address
281               .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "VolForPlotting" & CStr(i)).Address
282           End With

              Dim xValuesRange As Range
283           With RangeFromSheet(ws, "ScenarioDefinition" & CStr(i))
284               Set xValuesRange = .Cells(1, cnd_Months).Resize(.Rows.Count)
285           End With

286           With cht.SeriesCollection.NewSeries
287               .ChartType = xlXYScatterLinesNoMarkers
288               .Name = L_HedgeCapacity
289               .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
290               .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "HedgeCapacity" & CStr(i)).Address
291           End With

292           With cht.SeriesCollection.NewSeries
293               .ChartType = xlXYScatterLinesNoMarkers
294               .Name = L_LineEx
295               .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
296               .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "LineExhaustionLevel" & CStr(i)).Address
297           End With

298           With cht.SeriesCollection.NewSeries
299               .ChartType = xlXYScatterLinesNoMarkers
300               .Name = L_HedgeCompetion
301               .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
302               .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "HedgeCompletionRatio" & CStr(i)).Address
303           End With

304           cht.SeriesCollection(1).AxisGroup = 1
305           cht.SeriesCollection(2).AxisGroup = 1
306           cht.SeriesCollection(3).AxisGroup = 2
307           cht.SeriesCollection(4).AxisGroup = 2
308           cht.SeriesCollection(5).AxisGroup = 2

309           cht.Axes(xlCategory).MinimumScale = 0
310           cht.Axes(xlCategory).MaximumScale = Choose(i, MaxMonth1, MaxMonth2)
311           cht.Axes(xlCategory).MajorUnit = 6
312           cht.Axes(xlValue, xlPrimary).MinimumScale = MinimumScale
313           cht.Axes(xlValue, xlPrimary).MaximumScale = MaximumScale

              'Have to change the ChartType late in the code...
314           cht.SeriesCollection(3).ChartType = xlColumnClustered
315           cht.SetElement (msoElementLegendBottom)
316           cht.SetElement (msoElementChartTitleAboveChart)
317           With cht.ChartTitle.Format.TextFrame2.TextRange.Font
318               .Bold = msoFalse
319               .Size = 14
320               .Italic = msoFalse
321           End With
322           cht.ChartTitle.Caption = "=ScenarioCompare!ScenarioDescription" & CStr(i)

323           With cht.Legend
324               .Left = 10
325               .Width = 570
326               .Height = 43.748
327               .Top = 324.651
328           End With
329           cht.PlotArea.Height = 290

              'This is strange. I have seen the chart end up with many (15!) series, but cannot replicate that mis-behaviour
              Dim j As Long
330           For j = cht.SeriesCollection.Count To 6 Step -1
331               cht.SeriesCollection(j).Delete
332           Next j

333       Next i

334       ColorDifferences ws.Range("ScenarioDefinition1"), ws.Range("ScenarioDefinition2")

          Dim KPIDiffs
335       KPIDiffs = sArrayRange(KPIs1, sSubArray(KPIs2, 1, 2, , 1), sSubArray(sArraySubtract(KPIs2, KPIs1), 1, 2, , 1))
336       KPIDiffs(1, 2) = "Scenario 1": KPIDiffs(1, 3) = "Scenario 2": KPIDiffs(1, 4) = "Scen 2 minus Scen 1"

337       With ws.Range(KPIAddressDiffs).Resize(sNRows(KPIDiffs), sNCols(KPIDiffs))
338           .Value = sArrayExcelString(KPIDiffs)
339           .Cells(1, 1).Font.Bold = True
340           .Columns.AutoFit
341           .NumberFormat = "0.0"
342           .Rows(2).NumberFormat = "General"
343           AddGreyBorders .offset(0), True
344           ws.Names.Add "KPIDiffs", .offset(0)
345           .offset(0, 1).Resize(, .Columns.Count - 1).HorizontalAlignment = xlHAlignCenter
346       End With

          Dim Differences

347       Differences = sArrayRange(sIntegers(SafeMax(NR1, NR2)), HedgeCapacity1, HedgeCapacity2, _
              sArraySubtract(HedgeCapacity2, HedgeCapacity1))
348       Differences = sArrayStack(sArrayRange("Month", "Hedge Capacity 1 ($bln Left axis)", _
              "Hedge Capacity 2 ($bln Left axis)", "2 minus 1 ($bln Right axis)"), Differences)

349       Set TargetRange = TargetRange.Resize(sNRows(Differences), sNCols(Differences))

          Dim SourceRange As Range
350       Set SourceRange = TargetRange

351       With TargetRange
352           .Value = Differences
353           AddGreyBorders .offset(0), True
354           ws.Names.Add "HedgeCapacityDifferences", .offset(0)
355           .NumberFormat = "0"
356           .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
357           Set TargetRange = .Cells(1, .Columns.Count + 2)
358       End With

359       Set TL = ws.Range(TopLeftAddress3)

360       Set cht = ws.ChartObjects.Add(Left:=TL.Left, Top:=TL.Top, Width:=590, Height:=374.4).Chart
361       cht.ChartType = xlXYScatterLinesNoMarkers
362       cht.Parent.Placement = xlMove

363       With cht.SeriesCollection.NewSeries    'Hedge Capacity 1
364           .ChartType = xlColumnClustered
365           .Name = SourceRange.Cells(1, 2).Value
366           .xValues = "='" & ws.Name & "'!" & SourceRange.offset(1, 0).Resize(SourceRange.Rows.Count - 1, 1).Address
367           .Values = "='" & ws.Name & "'!" & SourceRange.offset(1, 1).Resize(SourceRange.Rows.Count - 1, 1).Address
368       End With

369       With cht.SeriesCollection.NewSeries    'Hedge Capacity 2
370           .ChartType = xlColumnClustered
371           .Name = SourceRange.Cells(1, 3).Value
372           .xValues = "='" & ws.Name & "'!" & SourceRange.offset(1, 0).Resize(SourceRange.Rows.Count - 1, 1).Address
373           .Values = "='" & ws.Name & "'!" & SourceRange.offset(1, 2).Resize(SourceRange.Rows.Count - 1, 1).Address
374       End With

375       With cht.SeriesCollection.NewSeries    'Differences
376           .ChartType = xlXYScatterLinesNoMarkers
377           .Name = SourceRange.Cells(1, 4).Value
378           .xValues = "='" & ws.Name & "'!" & SourceRange.offset(1, 0).Resize(SourceRange.Rows.Count - 1, 1).Address
379           .Values = "='" & ws.Name & "'!" & SourceRange.offset(1, 3).Resize(SourceRange.Rows.Count - 1, 1).Address
380           .AxisGroup = 2
381       End With

382       cht.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0"
383       cht.SetElement (msoElementLegendBottom)
384       cht.SetElement (msoElementChartTitleAboveChart)
385       cht.ChartTitle.Text = "Hedge Capacity Impact"
386       With cht.ChartTitle.Format.TextFrame2.TextRange.Font
387           .Bold = msoFalse
388           .Size = 14
389           .Italic = msoFalse
390       End With

391       Exit Sub
ErrHandler:
392       Throw "#RefreshScenarioCompareSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ColorDifferences
' Author    : Philip Swannell
' Date      : 01-Mar-2017
' Purpose   : For two ranges colour yellow the cells in each that differ from the corresponding
'             cell in the other. If ranges are of different sizes the "extra" cells are also coloured yellow.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ColorDifferences(ByVal R1 As Range, ByVal R2 As Range)
          Dim c As Range
          Dim D As Range
          Dim NC As Long
          Dim NR As Long
          Dim R3 As Range
          Dim R4 As Range

1         On Error GoTo ErrHandler
2         NR = SafeMin(R1.Rows.Count, R2.Rows.Count)
3         NC = SafeMin(R1.Columns.Count, R2.Columns.Count)

4         Set R3 = R1.Resize(NR, NC)
5         Set R4 = R2.Resize(NR, NC)

6         For Each c In R3.Cells
7             Set D = R4.Cells(c.Row - R3.Row + 1, c.Column - R3.Column + 1)
8             If c.Value <> D.Value Then
9                 c.Interior.Color = 6750207
10                D.Interior.Color = 6750207
11            End If
12        Next c

13        If R3.Address <> R1.Address Then
14            IntersectWithComplement(R1, R3).Interior.Color = 6750207
15        End If

16        If R4.Address <> R2.Address Then
17            IntersectWithComplement(R2, R4).Interior.Color = 6750207
18        End If

19        Exit Sub
ErrHandler:
20        Throw "#ColorDifferences (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

