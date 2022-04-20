Attribute VB_Name = "modLoadSave"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modLoadSave
' Author    : Philip Swannell
' Date      : 16-Jun-2015
' Purpose   : Code for reading and writing scenario definitions and results to file
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuScenarioDefinitionSheet
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : linked to Menu button on Scenario Definition sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuScenarioDefinitionSheet()
          Dim Choices As Variant
          Dim chOpenOtherBooks As Variant
          Dim Enabled As Variant
          Dim enbOpenOtherBooks As Variant
          Dim FaceIDs As Variant
          Dim FidOpenOtherBooks As Variant
          Dim LinesBookIsOpen As Boolean
          Dim MarketBookIsOpen As Boolean
          Dim OBAO As Boolean
          Dim Res As Variant
          Dim TradesBookIsOpen As Boolean
          Const chCalc As String = "&Calculate                            (Shift F9)"
          Const chRunScenario As String = "&Run Scenario..."
          Const chRunMany As String = "Run &Many Scenarios..."
          Const chResize As String = "Change the &number of months in the scenario..."
          Const chDrillDown As String = "Speed &Grid drill-down"
          Const chOpenDefn As String = "&Open Scenario Definition"
          Const chSaveDefn As String = "&Save Scenario Definition"

1         On Error GoTo ErrHandler
2         RunThisAtTopOfCallStack

3         OBAO = OtherBooksAreOpen(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen)
4         If OBAO Then
5             chOpenOtherBooks = CreateMissing()
6             FidOpenOtherBooks = CreateMissing()
7             enbOpenOtherBooks = CreateMissing()
8         Else
9             chOpenOtherBooks = NameForOpenOthers(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen, False)
10            FidOpenOtherBooks = 23
11            enbOpenOtherBooks = True
12        End If

13        FaceIDs = sArrayStack(FidOpenOtherBooks, 283, 156, 157, 296, 25)
14        Choices = sArrayStack(chOpenOtherBooks, chCalc, "--" & chRunScenario, chRunMany, "--" & chResize, chDrillDown)
15        Enabled = sArrayStack(enbOpenOtherBooks, sReshape(OBAO, 5, 1))

          Dim MRUChoices
          Dim MRUEnableFlags
          Dim MRUFaceIDs
          Dim MRUFileNames
16        GetMRUList gRegKey_Defn, MRUFileNames, MRUChoices, MRUFaceIDs, MRUEnableFlags, False
17        If IsEmpty(MRUFileNames) Then
18            FaceIDs = sArrayStack(FaceIDs, 23, 3)
19            Choices = sArrayStack(Choices, "--" & chOpenDefn, chSaveDefn)
20            Enabled = sArrayStack(Enabled, OBAO, OBAO)
21        Else
22            FaceIDs = sArrayStack(FaceIDs, MRUFaceIDs, 3)
23            Choices = sArrayStack(Choices, sArrayRange("--" & chOpenDefn, MRUChoices), chSaveDefn)
24            Enabled = sArrayStack(Enabled, sReshape(OBAO, sNRows(MRUFaceIDs) + 1, 1))
25        End If

26        Res = ShowCommandBarPopup(Choices, FaceIDs, Enabled, , ChooseAnchorObject())

27        If Res = "#Cancel!" Then Exit Sub

28        JuliaLaunchForCayley

29        Select Case Res
              Case Unembellish(CStr(chOpenOtherBooks))
30                OpenOtherBooks
31            Case Unembellish(chCalc)
32                RefreshScenarioDefinition True, False
33            Case Unembellish(chRunScenario)
34                OpenOtherBooks
35                BuildModelsInJulia False, 1, 1
36                RunScenario False, "", MN_CM, True, True, gModel_CM
37            Case Unembellish(chResize)
38                ResizeScenarioDefinition
39            Case Unembellish(chRunMany)
40                OpenOtherBooks
41                BuildModelsInJulia False, 1, 1
42                RunManyScenarios Empty, False, MN_CM, gModel_CM, RangeFromSheet(shConfig, "ScenarioResultsDirectory").Value
43            Case Unembellish(chOpenDefn)
44                OpenScenarioFile "", True, True
45            Case Unembellish(chSaveDefn)
46                RefreshScenarioDefinition True, False
47                SaveScenarioDefinitionFile "", RangeFromSheet(shScenarioDefinition, "ScenarioDescription")
48            Case Unembellish(chDrillDown)
49                RefreshScenarioDefinition False, True
50            Case Else
                  Dim WhichFile
51                WhichFile = sStringBetweenStrings(Res, , " ")
52                If IsNumeric(WhichFile) Then
53                    WhichFile = MRUFileNames(CLng(WhichFile), 1)
54                    OpenScenarioFile CStr(WhichFile), True, True
55                Else
56                    OpenScenarioFile "", True, True
57                End If
58        End Select
59        Exit Sub
ErrHandler:
60        SomethingWentWrong "#MenuScenarioDefinitionSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuScenarioResultsSheet
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : linked to Menu button on Scenario Results sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuScenarioResultsSheet()
          Dim Choices
          Dim FaceIDs
          Dim Res
          Const chOpenResults = "&Open Scenario Results"
          Const chSaveResults = "&Save Scenario Results"
          Const chPasteCharts = "&Paste Charts to new workbook..."

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

          Dim MRUChoices
          Dim MRUEnableFlags
          Dim MRUFaceIDs
          Dim MRUFileNames
3         GetMRUList gRegKey_Res, MRUFileNames, MRUChoices, MRUFaceIDs, MRUEnableFlags, False
4         If IsEmpty(MRUFileNames) Then
5             FaceIDs = sArrayStack(23, 3, 422)
6             Choices = sArrayStack(chOpenResults, chSaveResults, "--" & chPasteCharts)
7         Else
8             FaceIDs = sArrayStack(MRUFaceIDs, 3, 422)
9             Choices = sArrayStack(sArrayRange(chOpenResults, MRUChoices), chSaveResults, "--" & chPasteCharts)
10        End If

11        Res = ShowCommandBarPopup(Choices, FaceIDs, , , ChooseAnchorObject())

12        If Res = "#Cancel!" Then Exit Sub

13        Select Case Res
              Case Unembellish(chOpenResults)
14                OpenScenarioFile "", False, True
15            Case Unembellish(chSaveResults)
16                SaveScenarioResultsFile "", RangeFromSheet(shScenarioResults, "ScenarioDescription")
17            Case Unembellish(chPasteCharts)
18                PasteManyScenarioCharts
19            Case Else
                  Dim WhichFile
20                WhichFile = sStringBetweenStrings(Res, , " ")
21                If IsNumeric(WhichFile) Then
22                    WhichFile = MRUFileNames(CLng(WhichFile), 1)
23                    OpenScenarioFile CStr(WhichFile), False, True
24                Else
25                    OpenScenarioFile "", False, True
26                End If
27        End Select
28        Exit Sub
ErrHandler:
29        SomethingWentWrong "#MenuScenarioResultsSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenScenarioFile
' Author    : Philip Swannell
' Date      : 16-Jun-2015
' Purpose   : Read data from file and splat it to appropriate sheets of this workbook,
'             works for both Scenario Definition Files and Scenario Results Files
' -----------------------------------------------------------------------------------------------------------------------
Sub OpenScenarioFile(FileName As String, isDefinition As Boolean, WithLinesHistory As Boolean)
          Dim CopyOfErr As String
          Dim CurrentVersion
          Dim FileFilter As String
          Dim i As Long
          Dim IsDefinitionInFile As Boolean
          Dim Map As Range
          Dim oldBlockChange As Boolean
          Dim OverwriteWith As Variant
          Dim RangeName As String
          Dim RegKey As String
          Dim SheetName As String
          Dim SourceRange As Range
          Dim SourceSheet As Worksheet
          Dim SPH As Object
          Dim SPH2 As clsSheetProtectionHandler
          Dim SUH As Object
          Dim TargetRange As Range
          Dim TargetSheet As Worksheet
          Dim Title As String
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent
3         gBlockChangeEvent = True
4         CurrentVersion = RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value

5         If isDefinition Then
6             RegKey = gRegKey_Defn
7             FileFilter = "ScenarioDefinitionFile (*.sdf),*.sdf"
8             Title = "Open Scenario Definition File"
9         Else
10            RegKey = gRegKey_Res
11            FileFilter = "ScenarioResultsFile (*.srf),*.srf"
12            Title = "Open Scenario Results File"
13        End If

14        If FileName = "" Then
15            FileName = GetOpenFilenameWrap(RegKey, FileFilter, 1, Title, , False, False)
16            If FileName = "False" Then GoTo EarlyExit
17        End If

          'Gets added back to MRU at the end of the method if all is well with the file
18        RemoveFileFromMRU RegKey, FileName

19        Set SUH = CreateScreenUpdateHandler()

          'Close it if it's already open  (because of an earlier error?)
20        If IsInCollection(Application.Workbooks, sSplitPath(FileName)) Then
21            Application.Workbooks(sSplitPath(FileName)).Close False
22        End If

23        If Not sFileExists(FileName) Then
24            Throw "Sorry, we couldn't find '" & FileName & "'" & vbLf & _
                  "Is it possible it was moved, renamed or deleted?", True
25        End If

26        Set wb = Application.Workbooks.Open(FileName, , , , "Foo")
27        Set SourceSheet = wb.Worksheets(1)
28        If Not IsInCollection(SourceSheet.Names, "Map") Then
29            Throw "Unexpected error, cannot find 'Map' in file. " & _
                  "So file is not a valid Scenario Definition or Secenario Results file"
30        End If
31        Set Map = RangeFromSheet(SourceSheet, "Map")

          Dim SavedByVersion

32        SavedByVersion = sVLookup("SavedByVersion", Map.Value)
33        If Not IsNumber(SavedByVersion) Then
34            Throw FileName & " was saved by an old version of the Cayley workbook that used a file format " & _
                  "incompatible with the current version of the Cayley workbook. Sorry, but you cannot use it.", True
35        ElseIf SavedByVersion < gOldestSupportedScenarioVersion Then
36            Throw FileName & " was saved by version " & CStr(SavedByVersion) & " of the Cayley workbook. " & _
                  "Unfortunately, this version of the Cayley workbook (" & CStr(CurrentVersion) & _
                  ") is not compatible with files that old and you cannot open them.", True
37        End If

38        IsDefinitionInFile = sVLookup("IsDefinition", Map.Value)
39        If Not VarType(IsDefinitionInFile) = vbBoolean Then
40            Throw FileName & " was saved by an old version of the Cayley workbook that used a " & _
                  "file format incompatible with the current version of the Cayley workbook. Sorry"
41        ElseIf IsDefinitionInFile <> isDefinition Then
42            Throw "That file is a Scenario " & IIf(IsDefinitionInFile, "Definition", "Results") & _
                  " file, but you need to open a Scenario " & _
                  IIf(isDefinition, "Definition", "Results") & " file"
43        End If

          'Set the options strategy to defaults - if the file contains a _
           different options strategy then that will overwrite
44        Set SPH = CreateSheetProtectionHandler(shCreditUsage)

          'The FutureTrades sheet could be in a "cleared out" state, in which _
           case there is no range "TheTrades" - we need to have one for later code to work...
45        If Not IsInCollection(shFutureTrades.Names, "TheTrades") Then
46            If IsNumber(sMatch("TheTrades", Map.Columns(2).Value)) Then
47                shFutureTrades.Names.Add "TheTrades", RangeFromSheet(shFutureTrades, "Headers").offset(1)
48            End If
49        End If

50        For i = 1 To Map.Rows.Count - 2
51            SheetName = Map.Cells(i, 1).Value
52            RangeName = Map.Cells(i, 2).Value
53            MorphRangeAndSheetNames SheetName, OverwriteWith

54            If Not IsInCollection(ThisWorkbook.Worksheets, SheetName) Then
55                Throw "Cannot find sheet named " & SheetName
56            End If
57            Set TargetSheet = ThisWorkbook.Worksheets(SheetName)

58            If Not IsInCollection(TargetSheet.Names, RangeName) Then
59                Throw "Cannot find range named " & RangeName & " on sheet " & TargetSheet.Name
60            End If
61            Set TargetRange = RangeFromSheet(TargetSheet, RangeName)
62            Set SourceRange = RangeFromSheet(SourceSheet, Map.Cells(i, Map.Columns.Count))
63            If TargetRange.Columns.Count <> SourceRange.Columns.Count Then
64                Throw "Different number of columns in datafile versus this workbook for range name " & _
                      RangeName & " on worksheet " & TargetSheet.Name
65            End If
66        Next i

          Dim D As New Dictionary

67        For i = 1 To Map.Rows.Count - 2
68            SheetName = Map.Cells(i, 1).Value
69            RangeName = Map.Cells(i, 2).Value
70            MorphRangeAndSheetNames SheetName, OverwriteWith
71            If SheetName = shScenarioResults.Name Then
                  'Load into a Dictionary for subsequant call to RefreshScenarioResultsSheet
72                D.Add RangeName, SourceSheet.Range(Map(i, 3)).Value
73            Else

74                Set TargetSheet = ThisWorkbook.Worksheets(SheetName)
75                Set SPH = Nothing
76                Set SPH = CreateSheetProtectionHandler(TargetSheet)
77                Set TargetRange = RangeFromSheet(TargetSheet, RangeName)
78                Set SourceRange = RangeFromSheet(SourceSheet, Map.Cells(i, Map.Columns.Count))
79                TargetRange.ClearContents
80                With TargetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
81                    If IsEmpty(OverwriteWith) Then
82                        .Value2 = sArrayExcelString(SourceRange.Value2)
83                    Else
84                        .Value2 = OverwriteWith
85                    End If
86                    TargetSheet.Names.Add RangeName, .offset(0)
87                End With
88            End If
89        Next i

          'For backward compatibility with Cayley 2017 vintage
90        If Not D.Exists("HedgeHorizon") Then
91            D.Add "HedgeHorizon", (sNCols(D("ScenarioResultsHeaders")) - 1) / 2
92        End If

93        If Not isDefinition Then    'Ship in the LinesHistory...
94            If WithLinesHistory Then
95                Set SPH2 = CreateSheetProtectionHandler(shLinesHistory)
96                If Not IsInCollection(wb.Worksheets, "LinesHistory") Then
97                    Throw "Cannot find sheet 'LinesHistory' in file " & wb.FullName
98                End If

                  Dim N As Name
99                For Each N In shLinesHistory.Names
100                   N.Delete
101               Next
102               shLinesHistory.UsedRange.EntireColumn.Delete
103               wb.Worksheets("LinesHistory").UsedRange.Copy shLinesHistory.Cells(1, 1)
104               For Each N In wb.Worksheets("LinesHistory").Names
105                   shLinesHistory.Names.Add sStringBetweenStrings(N.Name, "!"), _
                          shLinesHistory.Range(N.RefersToRange.Address)
106               Next N
107               shLinesHistory.UsedRange.Columns.AutoFit
108           End If
109       End If

110       gBlockChangeEvent = oldBlockChange

111       If isDefinition Then
112           RefreshScenarioDefinition False
113           Set SPH = CreateSheetProtectionHandler(shScenarioDefinition)
114           RangeFromSheet(shScenarioDefinition, "FileName").Value = "'" & FileName
115       Else
116           RefreshScenarioResultsSheet shScenarioResults, _
                  DictGet(D, "ScenarioDefinitionHeaders"), DictGet(D, "ScenarioDefinition"), _
                  DictGet(D, "ScenarioResultsHeaders"), DictGet(D, "ScenarioResults"), _
                  FileName, DictGet(D, "ShocksDerivedFrom"), _
                  DictGet(D, "HistoryStart"), DictGet(D, "HistoryEnd"), _
                  DictGet(D, "BaseSpot"), DictGet(D, "BaseVol"), _
                  DictGet(D, "ForwardsRatio"), DictGet(D, "PutRatio"), _
                  DictGet(D, "CallRatio"), DictGet(D, "PutStrikeOffset"), _
                  DictGet(D, "CallStrikeOffset"), DictGet(D, "StrategySwitchPoints"), _
                  DictGet(D, "AllocationByYear"), DictGet(D, "ModelType"), _
                  DictGet(D, "NumMCPaths"), DictGet(D, "NumObservations"), _
                  DictGet(D, "FilterBy2"), DictGet(D, "Filter2Value"), _
                  DictGet(D, "IncludeAssetClasses"), DictGet(D, "CurrenciesToInclude"), _
                  DictGet(D, "TradesScaleFactor"), DictGet(D, "LinesScaleFactor"), _
                  DictGet(D, "TimeStart"), DictGet(D, "TimeEnd"), _
                  DictGet(D, "ComputerName"), DictGet(D, "UseSpeedGrid"), _
                  DictGet(D, "SpeedGridWidth"), DictGet(D, "HighFxSpeed"), DictGet(D, "LowFxSpeed"), _
                  DictGet(D, "VaryGridWidth"), DictGet(D, "SpeedGridBaseVol"), DictGet(D, "AnnualReplenishment"), _
                  DictGet(D, "ScenarioDescription"), DictGet(D, "HedgeHorizon")

117           FormatFutureTradesSheet

118       End If

119       wb.Close False
          'Add file to Most Recently Used list only if the opening was successful
120       AddFileToMRU RegKey, FileName

EarlyExit:
121       Set Map = Nothing
122       Set SourceSheet = Nothing
123       Set wb = Nothing
124       gBlockChangeEvent = oldBlockChange

125       Exit Sub
ErrHandler:
126       CopyOfErr = "#OpenScenarioFile (line " & CStr(Erl) & "): " & Err.Description & "!"
127       gBlockChangeEvent = oldBlockChange
128       If Not wb Is Nothing Then wb.Close False
129       Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MorphRangeAndSheetNames
' Author    : Philip Swannell
' Date      : 08-Sep-2016
' Purpose   : Scenario files saved out with earlier versions of the Cayley workbook need
'             some morphing...
'             Code can set OverwriteWith to a new value to be placed into the target range
'             or to Missing so that the target range is left unaltered
'             Following complete re-write of Scenario Code Oct 2016 we abandon trying
'             to support file formats pre-dating this time, so this method much simpler
'             than previously.
' -----------------------------------------------------------------------------------------------------------------------
Function MorphRangeAndSheetNames(ByRef SheetName As String, ByRef OverwriteWith As Variant) As String
1         On Error GoTo ErrHandler

2         OverwriteWith = Empty
3         If SheetName = "PFE" Then
4             SheetName = shCreditUsage.Name    'changed 21 Jan 2017
5         End If
6         Exit Function
ErrHandler:
7         Throw "#MorphRangeAndSheetNames (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatFutureTradesSheet
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Applies number formatting to range TheTrades in sheet FutureTrades
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatFutureTradesSheet()
          Dim i As Long
          Dim SPH As Object

1         On Error GoTo ErrHandler
2         If IsInCollection(shFutureTrades.Names, "TheTrades") Then
3             Set SPH = CreateSheetProtectionHandler(shFutureTrades)
4             With RangeFromSheet(shFutureTrades, "TheTrades")
5                 For i = 1 To .Columns.Count
6                     Select Case .Cells(0, i).Value
                          Case "FWD_1", "FWD_2"
7                             .Columns(i).NumberFormat = "#,##0;[Red]-#,##0"
8                         Case "DEAL_DATE", "MATURITY_DATE"
9                             .Columns(i).NumberFormat = "dd-mmm-yyyy"
10                    End Select
11                Next i
12            End With
13        End If
14        Exit Sub
ErrHandler:
15        Throw "#FormatFutureTradesSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PrepareScenarioDefinitionSheetForRelease
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : Put the sheet in a good state for releasing, but (for speed) don't require Julia to be running or
'             MarketData workbook to be open etc.
' -----------------------------------------------------------------------------------------------------------------------
Sub PrepareScenarioDefinitionSheetForRelease()
          Dim SPH As clsSheetProtectionHandler
          Dim ws As Worksheet
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler
          
          Dim AvEDSTraded As Variant
          Dim FxShocks As Variant
          Dim FxVolShocks As Variant
          Dim ReplenishmentAmounts
          Const ShocksDerivedFrom As String = "History"
          Dim HistoryEnd As Date
          Dim HistoryStart As Date
          Const ForwardsRatio = 1
          Const PutRatio = 0
          Const CallRatio = 0
          Const PutStrikeOffset = 0
          Const CallStrikeOffset = 0
          Const StrategySwitchPoints = Empty
          Const AllocationByYear = "0:0:1:1:1:0:0:0"
          Const UseSpeedGrid = False
          Const SpeedGridWidth = 0.05
          Const SpeedGridBaseVol = 0.0977
          Const HighFxSpeed = 0.5
          Const LowFxSpeed = 1.5
          Const VaryGridWidth = True
          Const AnnualReplenishment = 30000000000#
          Const TradesScaleFactor = 1
          Const LinesScaleFactor = 1
          Dim ScenarioDescription As String

2         Set ws = shScenarioDefinition
3         Set SPH = CreateSheetProtectionHandler(ws)
4         Set XSH = CreateExcelStateHandler(, , False)

5         HistoryStart = CDate("1-Jan-2000")
6         HistoryEnd = CDate("1-Jan-2002")

7         FxShocks = sArrayStack(0.965393794749403, 0.96887430389817, 0.952068416865553, 0.90950676213206, _
              0.923826571201273, 0.949980111376293, 0.917263325377884, 0.894689737470167, 0.878182179793158, _
              0.853420843277645, 0.869729514717582, 0.934268098647574, 0.936157517899761, 0.925715990453461, _
              0.874900556881464, 0.888126491646778, 0.841388225934765, 0.8425815433572, 0.874602227525855, _
              0.904932378679395, 0.912191726332538, 0.898667462211615, 0.891109785202864, 0.88494431185362)
          
8         FxVolShocks = sArrayStack(1.8088467614534, 1.90363349131122, 1.88783570300158, 2.0695102685624, _
              2.08135860979463, 1.97472353870458, 2.0695102685624, 2.17219589257504, 2.03791469194313, _
              2.25118483412322, 2.19194312796209, 2.20379146919431, 2.09320695102686, 2.01421800947867, _
              2.09715639810427, 1.99447077409163, 1.84834123222749, 1.7219589257504, 1.90521327014218, _
              1.91627172195893, 1.95892575039494, 1.85624012638231, 1.84044233807267, 1.80489731437599)
          
9         AvEDSTraded = sArrayStack(1.16908200931924, 1.13572770769406, 1.11308352703123, 1.09051751789976, _
              1.04766527965134, 1.09620433391191, 1.08404777247415, 1.04338567500259, 1.00498653824298, _
              0.98476539921892, 0.986338950965503, 1.03855514262984, 1.08285035280689, 1.06269607398568, _
              1.04784384356693, 1.03013139561314, 1.00953013904742, 0.984287089441982, 0.993572043827295, _
              1.04056103559199, 1.05208561455847, 1.04541259209298, 1.02487523323931, 1.0278619843164)

10        ReplenishmentAmounts = sReshape(AnnualReplenishment / 12, 24, 1)

11        ScenarioDescription = DescribeScenario(FxShocks, ReplenishmentAmounts, ShocksDerivedFrom, HistoryStart, _
              HistoryEnd, ForwardsRatio, PutRatio, CallRatio, PutStrikeOffset, CallStrikeOffset, StrategySwitchPoints, _
              AllocationByYear, UseSpeedGrid, SpeedGridWidth, HighFxSpeed, LowFxSpeed, VaryGridWidth, _
              AnnualReplenishment, TradesScaleFactor, LinesScaleFactor)

12        ResizeScenarioDefinition 24
13        RangeFromSheet(ws, "FileName").Value = "Not saved yet"
14        RangeFromSheet(ws, "ShocksDerivedFrom").Value = ShocksDerivedFrom
15        RangeFromSheet(ws, "HistoryStart").Value = HistoryStart
16        RangeFromSheet(ws, "HistoryEnd").Value = HistoryEnd
17        RangeFromSheet(ws, "ForwardsRatio").Value = ForwardsRatio
18        RangeFromSheet(ws, "PutRatio").Value = PutRatio
19        RangeFromSheet(ws, "CallRatio").Value = CallRatio
20        RangeFromSheet(ws, "PutStrikeOffset").Value = PutStrikeOffset
21        RangeFromSheet(ws, "CallStrikeOffset").Value = CallStrikeOffset
22        RangeFromSheet(ws, "StrategySwitchPoints").Value = StrategySwitchPoints
23        RangeFromSheet(ws, "AllocationByYear").Value = AllocationByYear
24        RangeFromSheet(ws, "UseSpeedGrid").Value = UseSpeedGrid
25        RangeFromSheet(ws, "SpeedGridWidth").Value = SpeedGridWidth
26        RangeFromSheet(ws, "HighFxSpeed").Value = HighFxSpeed
27        RangeFromSheet(ws, "LowFxSpeed").Value = LowFxSpeed
28        RangeFromSheet(ws, "VaryGridWidth").Value = VaryGridWidth
29        RangeFromSheet(ws, "SpeedGridBaseVol").Value = SpeedGridBaseVol
30        RangeFromSheet(ws, "AnnualReplenishment").Value = AnnualReplenishment
31        RangeFromSheet(ws, "ScenarioDescription").Value = ScenarioDescription

32        With RangeFromSheet(ws, "ScenarioDefinition")
33            With .offset(1).Resize(.Rows.Count - 1)
34                .Columns(1).Value = sIntegers(24)
                  'Enter "plausibe numbers", the FxShocks etc. get overwritten during interactive _
                   use of the ScenarioDefinition sheet.
35                .Columns(2).Value = FxShocks
36                .Columns(3).Value = FxVolShocks
37                .Columns(4).Value = ReplenishmentAmounts
38                .Columns(5).Value = AvEDSTraded
39            End With
40        End With

          ' RefreshScenarioDefinition True

          'Make one change on the ScenarioResults sheet...
41        Set ws = shScenarioResults
42        Set SPH = CreateSheetProtectionHandler(ws)
43        RangeFromSheet(ws, "FileName").Value = "Not yet saved to file"

44        Exit Sub
ErrHandler:
45        Throw "#PrepareScenarioDefinitionSheetForRelease (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeScenarioDefinition
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : Attached to menu button on the ScensraioDefinition sheet. Allows the user
'             to change the number of months in a Scenario.
' -----------------------------------------------------------------------------------------------------------------------
Sub ResizeScenarioDefinition(Optional ByVal NewNumRows As Long)
          Dim CurrentNumRows As Long
          Dim inputboxRes As Variant
          Dim SPH As Object
          Dim SUH As clsScreenUpdateHandler
          Const Title = "Change the number of months in the scenario"
          Dim c As Range
          Dim OldMaxMonth As Long
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         CurrentNumRows = RangeFromSheet(shScenarioDefinition, "ScenarioDefinition").Rows.Count - 1

3         If NewNumRows = 0 Then
TryAgain:
4             inputboxRes = InputBoxPlus("How many months should be in the Scenario?", Title, CStr(CurrentNumRows))
5             If VarType(inputboxRes) = vbBoolean Then Exit Sub
6             If Not IsNumeric(inputboxRes) Then GoTo TryAgain

7             On Error Resume Next
8             NewNumRows = CLng(inputboxRes)
9             On Error GoTo ErrHandler

10            If NewNumRows < 1 Then GoTo TryAgain
11            If NewNumRows > 120 Then GoTo TryAgain
12            If NewNumRows <> CLng(NewNumRows) Then GoTo TryAgain
13        End If

14        Set SPH = CreateSheetProtectionHandler(shScenarioDefinition)
15        Set SUH = CreateScreenUpdateHandler()

16        If NewNumRows > CurrentNumRows Then
17            With RangeFromSheet(shScenarioDefinition, "ScenarioDefinition")
18                OldMaxMonth = FirstElementOf(sMaxOfNums(.Columns(1)))
19                .ClearFormats
20                With .Resize(NewNumRows + 1)
21                    shScenarioDefinition.Names.Add "ScenarioDefinition", .offset(0)
22                    For Each c In .Columns(1).Cells
23                        If IsEmpty(c.Value) Then
24                            OldMaxMonth = OldMaxMonth + 1
25                            c.Value = OldMaxMonth
26                        End If
27                    Next c
28                End With
29            End With
30            RefreshScenarioDefinition False
31        ElseIf NewNumRows < CurrentNumRows Then
32            With RangeFromSheet(shScenarioDefinition, "ScenarioDefinition")
33                .offset(NewNumRows + 1).Resize(CurrentNumRows - NewNumRows).Clear
34                With .Resize(NewNumRows + 1)
35                    shScenarioDefinition.Names.Add "ScenarioDefinition", .offset(0)
36                End With
37                RefreshScenarioDefinition False
38            End With
39        Else
40            With RangeFromSheet(shScenarioDefinition, "ScenarioDefinition")
                  'some but not all of the cell formatting within RefreshScenarioDefinition
41                AddGreyBorders .offset(0), True
42                .HorizontalAlignment = xlHAlignCenter
43                .Rows(1).Font.Bold = True

44                With sColumnFromTable(.offset(0), "ReplenishmentAmount")
45                    .NumberFormat = "#,##0;[Red]-#,##0"
46                End With
47                With sColumnFromTable(.offset(0), "AvEDSTraded")
48                    .NumberFormat = "0.0000"
49                End With
50                With sColumnFromTable(.offset(0), "FxShock")
51                    .NumberFormat = "0.0000"
52                End With
53                With sColumnFromTable(.offset(0), "FxVolShock")
54                    .NumberFormat = "0.0000"
55                End With
56            End With
57        End If

58        Res = shScenarioDefinition.UsedRange.Rows.Count

59        Exit Sub
ErrHandler:
60        SomethingWentWrong "#ResizeScenarioDefinition (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveScenarioResultsFile
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : attached to menu button on ScenarioResults sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveScenarioResultsFile(FileName As String, ScenarioDescription As String)
1         On Error GoTo ErrHandler

          Dim InitialFileName As String

          Dim WhatToSaveArray As Variant
          
          'Save these from the ScenarioResults worksheet
          Const RangeNames = "ScenarioDefinitionHeaders,ScenarioDefinition,ScenarioResultsHeaders,ScenarioResults," & _
              "ShocksDerivedFrom,HistoryStart,HistoryEnd,BaseSpot,BaseVol,ForwardsRatio,PutRatio,CallRatio," & _
              "PutStrikeOffset,CallStrikeOffset,StrategySwitchPoints,AllocationByYear,ModelType,NumMCPaths," & _
              "NumObservations,FilterBy2,Filter2Value,IncludeAssetClasses,CurrenciesToInclude,TradesScaleFactor," & _
              "LinesScaleFactor,TimeStart,TimeEnd,ComputerName,UseSpeedGrid,SpeedGridWidth,HighFxSpeed,LowFxSpeed," & _
              "VaryGridWidth,SpeedGridBaseVol,AnnualReplenishment,ScenarioDescription"

2         If FileName = "" Then
3             InitialFileName = ScenarioDescriptionToFileName(ScenarioDescription, "srf")
4             FileName = GetSaveAsFilenameWrap(gRegKey_Res, InitialFileName, _
                  "ScenarioResultsFile (*.srf),*.srf", , "Save Scenario Definition File")
5             If FileName = "False" Then Exit Sub
6         End If

7         WhatToSaveArray = sTokeniseString(RangeNames)
8         WhatToSaveArray = sArrayRange(sReshape(shScenarioResults.Name, sNRows(WhatToSaveArray), 1), WhatToSaveArray)

          'Also save the HedgeHorizon from the Config sheet
9         WhatToSaveArray = sArrayStack(sArrayRange(shConfig.Name, "HedgeHorizon"), WhatToSaveArray)

10        WhatToSaveArray = sArrayStack(WhatToSaveArray, _
              sArrayRange(shFutureTrades.Name, "Headers"), sArrayRange(shFutureTrades.Name, "TheTrades"))

11        SaveDataFile WhatToSaveArray, FileName, False

12        AddFileToMRU gRegKey_Res, FileName
13        Exit Sub
ErrHandler:
14        Throw "#SaveScenarioResultsFile (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ScenarioDescriptionToFileName
' Author    : Philip Swannell
' Date      : 27-Oct-2016
' Purpose   : Encourage the use of descriptive file names
' -----------------------------------------------------------------------------------------------------------------------
Function ScenarioDescriptionToFileName(ScenarioDescription As String, ThreeLetterExtension As String)
          Dim FileName As String
          Dim i As Long
          Dim m As String

1         On Error GoTo ErrHandler
2         FileName = Replace(ScenarioDescription, vbCr, "")
3         FileName = Replace(FileName, vbLf, "")

          'Make dates use - delimiter
4         For i = 1 To 12
5             m = Format(CDate(CStr(i) & "/" & CStr(i) & "/2016"), "mmm")
6             FileName = Replace(FileName, m & " ", m & "-")
7         Next i
8         FileName = Replace(FileName, " - ", "-")

9         FileName = Replace(FileName, " ", "_")
          'replace characters illegal in file names
10        FileName = sRegExReplace(FileName, "\\|/|:|\*|\?|""|<|>|\|", "_")
          'replace consecutive underscores with a single underscore
11        FileName = sRegExReplace(FileName, "_{2,}", "_")
12        FileName = FileName & "." & LCase(ThreeLetterExtension)
13        ScenarioDescriptionToFileName = FileName
14        Exit Function
ErrHandler:
15        Throw "#ScenarioDescriptionToFileName (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveScenarioDefinitionFile
' Author    : Philip Swannell
' Date      : 16-Jul-2015
' Purpose   : This method returns an error string if it fails so that we can call via Application.Run with good
'             error handling.
' -----------------------------------------------------------------------------------------------------------------------
Function SaveScenarioDefinitionFile(FileName As String, ScenarioDescription As String)
1         On Error GoTo ErrHandler
          Dim InitialFileName As String
          Dim N As String
          Dim WhatToSaveArray

2         InitialFileName = ScenarioDescriptionToFileName(ScenarioDescription, "sdf")

3         N = shScenarioDefinition.Name

4         If FileName = "" Then
5             FileName = GetSaveAsFilenameWrap(gRegKey_Defn, InitialFileName, _
                  "ScenarioDefinitionFile (*.sdf),*.sdf", , "Save Scenario Definition File")
6             If FileName = "False" Then Exit Function
7         End If

8         RefreshScenarioDefinition True

9         WhatToSaveArray = sArrayStack(sArrayRange(N, "ShocksDerivedFrom"), _
              sArrayRange(N, "HistoryStart"), _
              sArrayRange(N, "HistoryEnd"), _
              sArrayRange(N, "ForwardsRatio"), _
              sArrayRange(N, "PutRatio"), _
              sArrayRange(N, "CallRatio"), _
              sArrayRange(N, "PutStrikeOffset"), _
              sArrayRange(N, "CallStrikeOffset"), _
              sArrayRange(N, "StrategySwitchPoints"), _
              sArrayRange(N, "AllocationByYear"), _
              sArrayRange(N, "ScenarioDefinition"), _
              sArrayRange(N, "UseSpeedGrid"), _
              sArrayRange(N, "SpeedGridWidth"), _
              sArrayRange(N, "HighFxSpeed"), _
              sArrayRange(N, "LowFxSpeed"), _
              sArrayRange(N, "VaryGridWidth"), _
              sArrayRange(N, "SpeedGridBaseVol"))

10        WhatToSaveArray = sArrayStack(WhatToSaveArray, _
              sArrayRange(N, "AnnualReplenishment"), _
              sArrayRange(shCreditUsage.Name, "ModelType"), _
              sArrayRange(shCreditUsage.Name, "NumMCPaths"), _
              sArrayRange(shCreditUsage.Name, "NumObservations"), _
              sArrayRange(shCreditUsage.Name, "FilterBy2"), _
              sArrayRange(shCreditUsage.Name, "Filter2Value"), _
              sArrayRange(shCreditUsage.Name, "IncludeAssetClasses"), _
              sArrayRange(shCreditUsage.Name, "TradesScaleFactor"), _
              sArrayRange(shCreditUsage.Name, "LinesScaleFactor"))

11        SaveDataFile WhatToSaveArray, FileName, True

12        AddFileToMRU gRegKey_Defn, FileName

13        Exit Function
ErrHandler:
14        Throw "#SaveScenarioDefinitionFile (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveDataFile
' Author    : Philip Swannell
' Date      : 16-Jun-2015
' Purpose   : Write the contents of arbitrary ranges from this workbook to a file that is
'             in fact a workbook. File created has a map for ease of reading back.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveDataFile(WhatToSaveArray, FileName As String, isDefinition As Boolean)
          Dim i As Long
          Dim R As Range
          Dim SourceRange As Range
          Dim SPH As Object
          Dim SUH As Object
          Dim TargetRange As Range
          Dim wb As Workbook
          Dim WriteRow As Long
          Dim ws As Worksheet

1         Set SUH = CreateScreenUpdateHandler()

2         On Error GoTo ErrHandler
3         For i = 1 To sNRows(WhatToSaveArray)
4             If Not IsInCollection(ThisWorkbook.Worksheets, CStr(WhatToSaveArray(i, 1))) Then
5                 Throw "Cannot find worksheet " & CStr(WhatToSaveArray(i, 1))
6             End If
7             Set ws = ThisWorkbook.Worksheets(CStr(WhatToSaveArray(i, 1)))
8             If Not IsInCollection(ws.Names, CStr(WhatToSaveArray(i, 2))) Then
9                 Throw "Cannot find range named " & CStr(WhatToSaveArray(i, 2)) & " on worksheet " & _
                      CStr(WhatToSaveArray(i, 1))
10            End If
11            Set R = Nothing
12            On Error Resume Next
13            Set R = ws.Names(CStr(WhatToSaveArray(i, 2))).RefersToRange
14            On Error GoTo ErrHandler
15            If R Is Nothing Then
16                Throw "Name " & CStr(WhatToSaveArray(i, 2)) & " on worksheet " & _
                      CStr(WhatToSaveArray(i, 1)) & " does not refer to a range"
17            End If
18            If R.Parent.Name <> WhatToSaveArray(i, 1) Then
19                Throw "Name " & CStr(WhatToSaveArray(i, 2)) & " on worksheet " & _
                      CStr(WhatToSaveArray(i, 1)) & " refers to a range on a different sheet"
20            End If
21        Next i

22        Set wb = Application.Workbooks.Add()
23        Set ws = wb.Worksheets(1)
24        ws.Name = "MappedData"

25        WhatToSaveArray = sArrayStack(WhatToSaveArray, _
              sArrayRange("SavedByVersion", RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value))
26        WhatToSaveArray = sArrayStack(WhatToSaveArray, _
              sArrayRange("IsDefinition", isDefinition))

27        With ws.Cells(1, 1).Resize(sNRows(WhatToSaveArray), sNCols(WhatToSaveArray) + 1)
28            .Value = WhatToSaveArray
29            ws.Names.Add ("Map"), .offset(0)
30            AddGreyBorders .offset(0), True
31            WriteRow = .Rows.Count + 2
32        End With

33        For i = 1 To sNRows(WhatToSaveArray) - 2        ' -1 since last two rows are version info and IsDefinition
34            Set SourceRange = ThisWorkbook.Worksheets(WhatToSaveArray(i, 1)).Range(WhatToSaveArray(i, 2))
35            Set TargetRange = ws.Cells(WriteRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
36            TargetRange.Value = sArrayExcelString(SourceRange.Value)
37            WriteRow = WriteRow + TargetRange.Rows.Count + 1
38            With RangeFromSheet(ws, "Map")
39                .Cells(i, .Columns.Count).Value = TargetRange.Address
40            End With
41        Next i

42        If Not isDefinition Then
43            shLinesHistory.Copy After:=wb.Sheets(1)
44        End If

45        Application.DisplayAlerts = False
46        wb.SaveAs FileName, xlOpenXMLWorkbook
47        wb.Close False

48        If isDefinition Then
49            Set SPH = CreateSheetProtectionHandler(shScenarioDefinition)
50            RangeFromSheet(shScenarioDefinition, "FileName") = "'" & FileName
51        Else
52            Set SPH = CreateSheetProtectionHandler(shScenarioResults)
53            RangeFromSheet(shScenarioResults, "FileName") = "'" & FileName
54        End If

55        Set ws = Nothing

56        Exit Sub
ErrHandler:
57        Throw "#SaveDataFile (line " & CStr(Erl) & "): " & Err.Description & " FileName = " & FileName
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunManyScenarios
' Author    : Philip Swannell
' Date      : 16-Jun-2015
' Purpose   : Run many scenarios, loading from file then writing results to file
' -----------------------------------------------------------------------------------------------------------------------
Sub RunManyScenarios(FileList As Variant, SilentMode As Boolean, ModelName As String, ModelBareBones As Dictionary, OutputDirectory As String)
          Dim Failures As Variant
          Dim i As Long
          Dim InputDirectory As String
          Dim NumFailed As Long
          Dim NumScenarios As Long
          Dim Prompt As String
          Dim Res
          Dim ResultsFile As String
          Dim StatusBarPrePrefix As String
          Dim ThisFileName As String

1         On Error GoTo ErrHandler

2         If IsEmpty(FileList) Or IsMissing(FileList) Then
3             FileList = GetOpenFilenameWrap(gRegKey_Defn, _
                  "ScenarioDefinitionFile (*.sdf),*.sdf", , _
                  "Open the Scenario Definitions to Run", , True)
4             If VarType(FileList) = vbBoolean Then Exit Sub
5         End If
6         Force2DArray FileList
7         NumScenarios = sNCols(FileList)

8         If Not SilentMode Then
9             Prompt = ""
10            For i = 1 To sNCols(FileList)
11                Prompt = Prompt & vbLf & CStr(i) & ")  " & sSplitPath(CStr(FileList(1, i)))
12            Next i
13            Prompt = "Run " & CStr(NumScenarios) & " Scenario" & IIf(NumScenarios > 1, "s", "") & vbLf & _
                  "Results will be written to " & OutputDirectory & vbLf & vbLf & _
                  "Scenario file names are:" & _
                  Prompt

14            If MsgBoxPlus(Prompt, _
                  vbOKCancel + vbQuestion + vbDefaultButton2, _
                  "Run Many Scenarios", , , , , 500) <> vbOK Then Exit Sub
15        End If

16        ThrowIfError sCreateFolder(OutputDirectory)

17        For i = 1 To NumScenarios
18            ThisFileName = FileList(1, i)
19            InputDirectory = sSplitPath(ThisFileName, False)
20            ResultsFile = sJoinPath(OutputDirectory, Replace(sSplitPath(ThisFileName, True), ".sdf", ".srf"))

21            OpenScenarioFile ThisFileName, True, False
22            StatusBarPrePrefix = "RunManyScenarios Scenario " & CStr(i) & " of " & CStr(NumScenarios) & " "
23            Res = RunScenario(True, StatusBarPrePrefix, ModelName, False, False, ModelBareBones)
24            If Not Left(Res, 1) = "#" Then
25                SaveScenarioResultsFile ResultsFile, ""
26                ThrowIfError sFileCopy(ResultsFile, sJoinPath(InputDirectory, sSplitPath(ResultsFile)))
27                ThrowIfError sFileCopy(ThisFileName, sJoinPath(OutputDirectory, sSplitPath(ThisFileName)))
28            Else
29                If NumFailed = 0 Then
30                    Failures = sArrayStack("Scenario " & CStr(i), sSplitPath(ThisFileName), Res)
31                Else
32                    Failures = sArrayStack(Failures, "", "Scenario " & CStr(i), sSplitPath(ThisFileName), Res)
33                End If
34                NumFailed = NumFailed + 1
35            End If
36        Next i

37        If NumFailed > 0 Then
38            Prompt = CStr(NumFailed) & " of the " & CStr(NumScenarios) & " scenarios failed." & vbLf & _
                  sConcatenateStrings(Failures, vbLf)
39            If gDebugMode Then Debug.Print Prompt
40            MsgBoxPlus Prompt, vbCritical, "Run Many Scenarios", , , , , 600
41        End If

42        Exit Sub
ErrHandler:
43        SomethingWentWrong "#RunManyScenarios (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteManyScenarioCharts
' Author    : Philip Swannell
' Date      : 23-Jun-2015
' Purpose   : Select a number N Scenario Results files and paste N versions of the graph
'             on the Scenario sheet to a new workbook
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteManyScenarioCharts()
          Dim FileName As String
          Dim FileNames As Variant
          Dim i As Long
          Dim N As Long
          Dim SourceRange As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As Object
          Dim TargetCell As Range
          Dim TargetCell2 As Range
          Dim TargetRange As Range
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         FileNames = GetOpenFilenameWrap(gRegKey_Res, "ScenarioResultsFile (*.srf),*.srf", 1, _
              "Paste Charts to new workbook. Please choose result files.", , True)
3         If VarType(FileNames) = vbBoolean Then Exit Sub

4         Set SUH = CreateScreenUpdateHandler()
5         Set SPH = CreateSheetProtectionHandler(shScenarioResults)

6         Force2DArray FileNames
7         FileNames = sArrayTranspose(FileNames)
8         FileNames = sSortedArray(FileNames)
9         N = sNRows(FileNames)

10        Set wb = Application.Workbooks.Add
11        Set TargetCell = wb.Worksheets(1).Cells(4, 1)
12        Set TargetCell2 = wb.Worksheets(1).Cells(4, 23)

13        For i = LBound(FileNames) To UBound(FileNames)
14            StatusBarWrap "Pasting chart " & CStr(i) & " of " & CStr(N)
15            FileName = FileNames(i, 1)
16            With TargetCell.offset(-1)
17                .Value = "'" & FileName
18                .Font.Color = g_Col_GreyText
19            End With
20            OpenScenarioFile FileName, False, i = UBound(FileNames)
21            CopyChart shScenarioResults.ChartObjects(1), TargetCell
22            Set TargetCell = TargetCell.offset(3 + shScenarioResults.ChartObjects(1).Chart.ChartArea.Height / _
                  TargetCell.Height)

              'Paste over the scenario definitions and KPIs

23            Set SourceRange = RangeFromSheet(shScenarioResults, "KPIs")

24            Set TargetRange = TargetCell2.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
25            SourceRange.Copy
26            TargetRange.PasteSpecial xlPasteAll
27            TargetRange.PasteSpecial xlPasteColumnWidths

28            Set SourceRange = Range(RangeFromSheet(shScenarioResults, "FileName").Cells(0, 1), _
                  RangeFromSheet(shScenarioResults, "ScenarioDescription"))

29            Set TargetRange = TargetCell2.offset(RangeFromSheet(shScenarioResults, "KPIs").Rows.Count)
30            Set TargetRange = TargetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
31            SourceRange.Copy
32            TargetRange.PasteSpecial xlPasteAll
33            TargetRange.PasteSpecial xlPasteColumnWidths
34            If i > LBound(FileNames) Then
35                TargetRange.ClearComments
36            End If

37            Set TargetCell2 = TargetCell2.offset(, 2)

38        Next i
39        StatusBarWrap False

40        With TargetCell.Parent.Cells(1, 1)
41            .Value = "Scenario Charts"
42            .Font.Size = 22
43        End With

44        TargetCell.Parent.Name = "ScenarioCharts"
45        Application.GoTo shScenarioResults.Cells(1, 1)
46        Application.GoTo TargetCell.Parent.Cells(1, 1)
47        ActiveWindow.DisplayGridlines = False
48        ActiveWindow.DisplayHeadings = False

49        Exit Sub
ErrHandler:
50        SomethingWentWrong "#PasteManyScenarioCharts (line " & CStr(Erl) & "): " & Err.Description & "!"
51        StatusBarWrap False
End Sub

