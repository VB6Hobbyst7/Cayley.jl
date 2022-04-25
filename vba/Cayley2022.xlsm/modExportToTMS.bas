Attribute VB_Name = "modExportToTMS"
Option Explicit
Const m_MsgBoxTitle = "Export to TMS"
Const m_RangesInRegistry = "WhereToExport,FeedRates,ExportTrades,ExportMarketData,ExportTable,ExportCharts,Scenarios"
Const m_RegKey = "Cayley2022"
Const m_RegSection = "ExportToTMSSetings"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunExport
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : The Main of this module. Attached to the "Export..." button on the ExportToTMS worksheet
' -----------------------------------------------------------------------------------------------------------------------
Sub RunExport()
          
          Const DoFxHeadroom As Boolean = True
          Const DoTradeHeadroom As Boolean = True
          Const ExportJPG As Boolean = True
          Dim AllBanks As Variant
          Dim RunDate As Date
          Dim ChooseVector
          Dim DoExportCharts As Boolean
          Dim DoExportMarketData As Boolean
          Dim DoExportTable As Boolean
          Dim DoExportTrades As Boolean
          Dim DoFeedRates As Boolean
          Dim DoScenarios As Boolean
          Dim FileNameTable As String
          Dim i As Long
          Dim LinesBook As Workbook
          Dim MarketDataWorkbook As Workbook
          Dim NumScenarios As Long
          Dim ScenarioList As Variant
          Dim SolumDotOut As String
          Dim TargetFolder As String
          Dim TimeEnd
          Dim TimeStart

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         OpenOtherBooks

4         If TradesWorkbookIsOutOfDate() Then
5             LoadTradesFromTextFiles , , , True
6         End If

7         Set LinesBook = OpenLinesWorkbook(True, False)
8         Set MarketDataWorkbook = OpenMarketWorkbook(True, False)

9         DoFeedRates = RangeFromSheet(shExportToTMS, "FeedRates", False, False, True, False, False)
10        DoExportTrades = RangeFromSheet(shExportToTMS, "ExportTrades", False, False, True, False, False)
11        DoExportMarketData = RangeFromSheet(shExportToTMS, "ExportMarketData", False, False, True, False, False)
12        DoExportTable = RangeFromSheet(shExportToTMS, "ExportTable", False, False, True, False, False)
13        DoExportCharts = RangeFromSheet(shExportToTMS, "ExportCharts", False, False, True, False, False)

14        If DoFeedRates Then
15            SolumDotOut = RangeFromSheet(MarketDataWorkbook.Worksheets("Config"), "MarketDataFile")
16            SolumDotOut = sJoinPath(MarketDataWorkbook.Path, SolumDotOut)
20        End If

          RunDate = Date

21        TargetFolder = RangeFromSheet(shExportToTMS, "WhereToExport", False, True, False, False, False)
22        TargetFolder = sJoinPath(TargetFolder, Format(RunDate, "yyyy-mm-dd"))
23        ThrowIfError sCreateFolder(TargetFolder)

          'Do very basic early validation of inputs for Scenarios. Do the SDF files exist, and is the list of them unique
24        ChooseVector = sArrayEquals(True, RangeFromSheet(shExportToTMS, "Scenarios").Columns(1).Value)
25        ChooseVector = sArrayAnd(ChooseVector, sArrayIsNonTrivialText(RangeFromSheet(shExportToTMS, "Scenarios").Columns(2).Value))
26        NumScenarios = sArrayCount(ChooseVector)
27        DoScenarios = NumScenarios >= 1
28        If DoScenarios Then
29            ScenarioList = sMChoose(RangeFromSheet(shExportToTMS, "Scenarios").Columns(2).Value, _
                  ChooseVector)
30            For i = 1 To sNRows(ScenarioList)
31                If Not sFileExists(ScenarioList(i, 1)) Then
32                    Throw "Cannot find the following scenario definition file" & vbLf & vbLf & ScenarioList(i, 1) & vbLf & vbLf & _
                          "Please ensure that all files given in cells " & Replace(RangeFromSheet(shExportToTMS, "Scenarios").Columns(2).Address, "$", "") & _
                          " are for sdf files that exist. You can double-click on those cells to browse.", True
33                ElseIf LCase(Right$(ScenarioList(i, 1), 4)) <> ".sdf" Then
34                    Throw "Scenario definition files should have .sdf file extension, but file " & vbLf & vbLf & ScenarioList(i, 1) & vbLf & "has a different extension." & vbLf & vbLf & _
                          "Please ensure that all files given in cells " & Replace(RangeFromSheet(shExportToTMS, "Scenarios").Columns(2).Address, "$", "") & _
                          " are for sdf files that exist. You can double-click on those cells to browse.", True
35                End If
36            Next i
              Dim Tmp
37            Tmp = sCountDistinctItems(ScenarioList)
38            If Tmp(1, 2) > 1 Then
39                Throw "Scenario file " & vbLf & Tmp(1, 1) & vbLf & "appears more than once in the list of scenarios to be executed." & vbLf & vbLf & _
                          "Please ensure that all files given in cells " & Replace(RangeFromSheet(shExportToTMS, "Scenarios").Columns(2).Address, "$", "") & _
                          " are for sdf files that exist and that they are all unique. You can double-click on those cells to browse.", True
40            End If
41        End If

42        If DoExportMarketData Then
              Dim MarketDataFolder As String
43            MarketDataFolder = ThrowIfError(sCreateFolder(sJoinPath(TargetFolder, "MarketData")))
44        End If
45        If DoExportTrades Then
              Dim TradesFolder As String
46            TradesFolder = ThrowIfError(sCreateFolder(sJoinPath(TargetFolder, "Trades")))
47        End If
48        If DoExportCharts Then
              Dim ChartsFolder As String
49            ChartsFolder = ThrowIfError(sCreateFolder(sJoinPath(TargetFolder, "Charts")))
50        End If
51        If DoExportTable Then
              Dim TableFolder As String
52            TableFolder = ThrowIfError(sCreateFolder(sJoinPath(TargetFolder, "Table")))
53        End If
54        If DoScenarios Then
              Dim ScenariosFolder As String
55            ScenariosFolder = ThrowIfError(sCreateFolder(sJoinPath(TargetFolder, "Scenarios")))
56        End If

57        If Not AreYouSure(DoFeedRates, DoExportTrades, DoExportMarketData, DoExportTable, DoExportCharts, DoScenarios, _
              MarketDataFolder, TradesFolder, ChartsFolder, TableFolder, ScenariosFolder, NumScenarios) Then Exit Sub

58        TimeStart = Now()
59        MessageLogWrite "Cayley RunExport START"

60        JuliaLaunchForCayley

61        ShowFileInSnakeTail , True
62        Application.StatusBar = "Export to TMS is running. Progress is shown in the SnakeTail application."

63        AllBanks = sSortedArray(GetColumnFromLinesBook("CPTY_PARENT", LinesBook))
64        PrepareForCalculation AllBanks(1, 1), True, True, True

65        If DoFeedRates Then
66            FeedRatesFromTextFile
67            BuildModelsInJulia True, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")
68        End If

69        If DoExportMarketData Then
70            ExportMarketData MarketDataFolder, RunDate
71        End If

72        If DoExportTrades Then
73            ExportTrades TradesFolder, RunDate
74        End If

75        If DoExportCharts Then
76            PasteCharts AllBanks, ChartsFolder, ExportJPG, RunDate, True
77        End If

78        If DoExportTable Then
79            RunTable DoTradeHeadroom, DoFxHeadroom, True, AllBanks
80            FileNameTable = sJoinPath(TableFolder, _
                  "ResultsByCounterpartyParent_" & Format(RunDate, "yyyy-mm-dd") & ".csv")
81            ExportTable FileNameTable
82        End If

83        If DoScenarios Then
84            BuildModelsInJulia False, 1, 1
85            RunManyScenarios ScenarioList, True, MN_CM, gModel_CM, ScenariosFolder
86        End If

87        shExportToTMS.Activate

88        TimeEnd = Now
89        SafeAppActivate shExportToTMS
90        MessageLogWrite "Cayley RunExport started at " & Format(TimeStart, "yyyy-mm-dd hh:mm:ss")
91        MessageLogWrite "Cayley RunExport ended at " & Format(TimeEnd, "yyyy-mm-dd hh:mm:ss")
92        MessageLogWrite "Cayley RunExport took " & Format(TimeEnd - TimeStart, "hh:mm:ss")

93        Exit Sub
ErrHandler:
94        SomethingWentWrong "#RunExport (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AreYouSure
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : Post an "Are you sure you want to do this?" dialog, with details of what will be executed.
' -----------------------------------------------------------------------------------------------------------------------
Function AreYouSure(DoFeedRates As Boolean, DoExportTrades As Boolean, DoExportMarketData As Boolean, _
          DoExportTable As Boolean, DoExportCharts As Boolean, DoScenarios As Boolean, _
          MarketDataFolder As String, TradesFolder As String, ChartsFolder As String, TableFolder As String, _
          ScenariosFolder As String, NumScenarios As Long) As Boolean

          Dim AnyTasks As Boolean
          Dim MarketWB As Workbook
          Dim Prompt As String

1         On Error GoTo ErrHandler

2         Set MarketWB = OpenMarketWorkbook(True, False)

3         Prompt = "Export files for Treasury Management System?" & vbLf & vbLf & _
              "The following tasks will be carried out:"
                   
4         If DoFeedRates Then
              Dim SolumDotOut As String
5             AnyTasks = True
6             SolumDotOut = RangeFromSheet(MarketWB.Worksheets("Config"), "MarketDataFile")
7             SolumDotOut = sJoinPath(MarketWB.Path, SolumDotOut)
                   
8             Prompt = Prompt & vbLf & vbLf & _
                  "Rates will be fed from file" & vbLf & _
                  SolumDotOut & vbLf & _
                  "to the market data workbook"
                   
9         End If
            
10        If DoExportTrades Then
11            AnyTasks = True
12            Prompt = Prompt & vbLf & vbLf & _
                  "Three trade data files will be copied to" & vbLf & _
                  TradesFolder
13        End If
              
14        If DoExportMarketData Then
15            AnyTasks = True
16            Prompt = Prompt & vbLf & vbLf & _
                  "Three market data files will be copied to" & vbLf & _
                  MarketDataFolder
17        End If
                  
18        If DoExportTable Then
19            AnyTasks = True
20            Prompt = Prompt & vbLf & vbLf & _
                  "The Table worksheet will be updated to record bank-by-bank trade headroom and fx headroom. Results will be saved to" & vbLf & _
                  TableFolder
21        End If

22        If DoExportCharts Then
23            AnyTasks = True
24            Prompt = Prompt & vbLf & vbLf & _
                  "Charts plotting the current PFE of each bank's trades with Airbus and the bank's lines to Airbus will be saved to" & vbLf & _
                  ChartsFolder
25        End If

26        If DoScenarios Then
27            AnyTasks = True
28            Prompt = Prompt & vbLf & vbLf & _
                  NumberToText(NumScenarios) & " scenario" & IIf(NumScenarios > 1, "s", "") & " will be run with scenario definition files (.sdf) and scenario result files (.srf) saved to:" & vbLf & _
                  ScenariosFolder
29        End If

30        If DoScenarios Or DoExportTable Then
31            AnyTasks = True
32            Prompt = Prompt & vbLf & vbLf & _
                  "The process may take some time. You can monitor progress in SnakeTail"
33        End If

34        If Not AnyTasks Then
35            Throw "Please select at least one task, from FeedRates through ExportCharts, or at least one scenario.", True
36        End If

37        AreYouSure = MsgBoxPlus(Prompt, vbQuestion + vbOKCancel, m_MsgBoxTitle) = vbOK
38        Exit Function
ErrHandler:
39        Throw "#AreYouSure (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function NumberToText(Number As Long)

1         On Error GoTo ErrHandler
2         If Number > 0 And Number <= 10 Then
3             NumberToText = Choose(Number, "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten")
4         Else
5             NumberToText = CStr(Number)
6         End If

7         Exit Function
ErrHandler:
8         Throw "#NumberToText (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Sub ExportTrades(TargetFolder As String, RunDate As Date)
          Dim Suffix As String

1         On Error GoTo ErrHandler

2         Suffix = "_" & Format(RunDate, "yyyy-mm-dd") & ".csv"
3         MessageLogWrite "Exporting trade data to " & TargetFolder

4         ThrowIfError sFileCopy(FileFromConfig("FxTradesCSVFile"), sJoinPath(TargetFolder, "FxTrades" & Suffix), True)
5         ThrowIfError sFileCopy(FileFromConfig("RatesTradesCSVFile"), sJoinPath(TargetFolder, "RatesTrades" & Suffix), True)
6         ThrowIfError sFileCopy(FileFromConfig("AmortisationCSVFile"), sJoinPath(TargetFolder, "Amortisation" & Suffix), True)

7         Exit Sub
ErrHandler:
8         Throw "#ExportTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveETMSToRegistry
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : Save the contents of the worksheet to the Windows Registry, called from the Change event
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveETMSToRegistry()
          Dim i As Long
          Dim RangeNames

1         On Error GoTo ErrHandler
2         RangeNames = sTokeniseString(m_RangesInRegistry)

3         For i = 1 To sNRows(RangeNames)
4             SaveSetting m_RegKey, m_RegSection, RangeNames(i, 1), sMakeArrayString(RangeFromSheet(shExportToTMS, CStr(RangeNames(i, 1))).Value2)
5         Next i

6         Exit Sub
ErrHandler:
7         Throw "#SaveETMSToRegistry (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetETMSFromRegistry
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : Restore the contents of the worksheet from the Windows registry. Called from the workbook open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub GetETMSFromRegistry()
          Dim i As Long
          Dim j As Long
          Dim RangeNames
          Dim Res
          Const MinNoScenarios = 20
          Const MaxNoScenarios = 40
          Dim UpdateThisOne As Boolean
          
          Dim SPH As clsSheetProtectionHandler
          Dim TargetRange As Range
          Dim XSH As clsExcelStateHandler
          
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shExportToTMS)
3         Set XSH = CreateExcelStateHandler(, , False) 'Disable events

4         RangeNames = sTokeniseString(m_RangesInRegistry)

5         For i = 1 To sNRows(RangeNames)
6             Res = GetSetting(m_RegKey, m_RegSection, RangeNames(i, 1), "#Not found!")
7             If Not sIsErrorString(Res) Then
8                 UpdateThisOne = True
9                 Res = sParseArrayString(CStr(Res))
10                Set TargetRange = RangeFromSheet(shExportToTMS, CStr(RangeNames(i, 1)))
11                If RangeNames(i, 1) = "Scenarios" Then
12                    If sNRows(Res) < MinNoScenarios Then
13                        Res = sArrayStack(Res, sReshape(Empty, MinNoScenarios - sNRows(Res), 2))
14                    ElseIf sNRows(Res) > MaxNoScenarios Then
15                        Res = sSubArray(Res, 1, 1, MaxNoScenarios)
16                    End If
17                    If sNCols(Res) <> 2 Then
18                        UpdateThisOne = False
19                    End If
20                    For j = 1 To sNRows(Res)
21                        If VarType(Res(j, 1)) <> vbBoolean Then
22                            Res(j, 1) = False
23                        End If
24                    Next
25                Else
26                    If sNCols(Res) <> 1 Or sNRows(Res) <> 1 Then
27                        UpdateThisOne = False
28                    End If
29                End If
30                If UpdateThisOne Then
31                    If TargetRange.Rows.Count = sNRows(Res) And TargetRange.Columns.Count = sNCols(Res) Then
32                        TargetRange.Value = sArrayExcelString(Res)
33                    Else
34                        TargetRange.ClearContents
35                        With TargetRange.Resize(sNRows(Res), sNCols(Res))
36                            .Value = Res
37                            shExportToTMS.Names.Add RangeNames(i, 1), .Cells
38                        End With
39                    End If
40                End If
41            End If
42        Next i

43        FormatExportToTMS

44        Exit Sub
ErrHandler:
45        Throw "#GetETMSFromRegistry (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FormatExportToTMS
' Author     : Philip Swannell
' Date       : 21-Mar-2022
' Purpose    : Applies cell formatting to the worksheet ExportToTMS
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatExportToTMS()
          Dim c As Range
          Dim SPH As clsSheetProtectionHandler
          Dim TopLeftAddress As String
          Dim TopRightAddress As String
          Dim ws As Worksheet
          
1         On Error GoTo ErrHandler
2         Set ws = shExportToTMS

3         Set SPH = CreateSheetProtectionHandler(ws)

4         Set c = RangeFromSheet(ws, "WhereToExport")
5         With c
6             CayleyFormatAsInput .Cells
7             With .Validation
8                 .Delete
9                 .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="=ISTEXT(" & c.Address & ")"
10                .IgnoreBlank = True
11                .InputTitle = c.offset(0, -1).Value
12                .ErrorTitle = "Export to TMS"
13                .InputMessage = "Double-click to browse"
14                .ErrorMessage = c.offset(0, -1).Value & " must be text"
15                .ShowInput = True
16                .ShowError = True
17            End With
18        End With

19        With Range(RangeFromSheet(ws, "FeedRates"), _
              RangeFromSheet(ws, "ExportCharts"))
20            TopLeftAddress = Replace(.Cells(1, 1).Address, "$", "")
21            CayleyFormatAsInput .Cells
22            .HorizontalAlignment = xlHAlignLeft
23            .FormatConditions.Delete
24            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & TopLeftAddress & "<>TRUE"
25            .FormatConditions(.FormatConditions.Count).SetFirstPriority
26            With .FormatConditions(1).Font
27                .ThemeColor = xlThemeColorDark1
28                .TintAndShade = -0.249946592608417
29            End With
30            .FormatConditions(1).StopIfTrue = False
              
31            For Each c In .Cells
32                With c
33                    With .Validation
34                        .Delete
35                        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                              xlBetween, Formula1:="=ISLOGICAL(" & c.Address & ")"
36                        .IgnoreBlank = True
37                        .InputTitle = c.offset(0, -1).Value
38                        .ErrorTitle = "Export to TMS"
39                        .InputMessage = "Double-click to change"
40                        .ErrorMessage = c.offset(0, -1).Value & " must be TRUE or FALSE"
41                        .ShowInput = True
42                        .ShowError = True
43                    End With
44                    If IsEmpty(c.Value) Then c.Value = False
45                End With
46            Next

47        End With

48        With RangeFromSheet(ws, "Scenarios")
49            TopLeftAddress = Replace(.Cells(1, 1).Address, "$", "")
50            TopRightAddress = Replace(.Cells(1, 2).Address, "$", "")
51            AddGreyBorders .Cells, True
52            CayleyFormatAsInput .Cells
53            .Columns(1).HorizontalAlignment = xlHAlignCenter
54            .Columns(2).HorizontalAlignment = xlHAlignLeft
55            .FormatConditions.Delete
56            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & "$" & TopLeftAddress & "<>TRUE"
57            .FormatConditions(.FormatConditions.Count).SetFirstPriority
58            With .FormatConditions(1).Font
59                .ThemeColor = xlThemeColorDark1
60                .TintAndShade = -0.249946592608417
61            End With
62            .FormatConditions(1).StopIfTrue = False
              
63            With .Columns(1).Validation
64                .Delete
65                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="=ISLOGICAL(" & TopLeftAddress & ")"
66                .IgnoreBlank = True
67                .InCellDropdown = True
68                .InputTitle = "DoThisOne?"
69                .ErrorTitle = "Export to TMS"
70                .InputMessage = "Double-click to change"
71                .ErrorMessage = "Value must be TRUE or FALSE"
72                .ShowInput = True
73                .ShowError = True
74            End With

75            With .Columns(2).Validation
76                .Delete
77                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="=ISTEXT(" & TopRightAddress & ")"
78                .IgnoreBlank = True
79                .InCellDropdown = True
80                .InputTitle = "ScenarioDefinitionFile"
81                .ErrorTitle = "Export to TMS"
82                .InputMessage = "Double-click to browse"
83                .ErrorMessage = "ScenarioDefinitionFile must be text"
84                .ShowInput = True
85                .ShowError = True
86            End With
87        End With
88        Exit Sub
ErrHandler:
91        Throw "#FormatExportToTMS (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExportMarketData
' Author     : Philip Swannell
' Date       : 21-Mar-2022
' Purpose    : Exports market data: Three files: 1) The current Solum.out, created by Airbus tech department, its location
'              is taken from the Config sheet of the Market Data Workbook. 2) & 3) the .json files that are parsed by the Julia code
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ExportMarketData(Folder As String, RunDate As Date)

          Dim AllCcys
          Dim CCysToBuild
          Dim FileName As String
          Dim FileNameH As String
          Dim MarketWB As Workbook
          Dim Numeraire As String

1         On Error GoTo ErrHandler
2         Set MarketWB = OpenMarketWorkbook(True, False)

          Dim SolumDotOut As String
3         SolumDotOut = RangeFromSheet(MarketWB.Worksheets("Config"), "MarketDataFile")
4         SolumDotOut = sJoinPath(MarketWB.Path, SolumDotOut)
5         If Not sFileExists(SolumDotOut) Then
6             Throw "Cannot find file '" & SolumDotOut & "' Please check the 'MarketDataFile' setting on the 'Config' worksheet of the MarketDataWorkbook"
7         End If

8         ThrowIfError sFileCopy(SolumDotOut, sJoinPath(Folder, "MarketDataFile_" & Format(RunDate, "yyyy-mm-dd") & ".out"), True)

9         FileName = sJoinPath(Folder, "CayleyMarket_" & Format(RunDate, "yyyy-mm-dd") & ".json")
10        FileNameH = sJoinPath(Folder, "CayleyMarketHistoricFxVol_" & Format(RunDate, "yyyy-mm-dd") & ".json")
11        Numeraire = NumeraireFromMDWB()
          '
12        AllCcys = AllCurrenciesInTradesWorkBook(True)
13        If LCase(RangeFromSheet(shConfig, "CurrenciesToInclude")) = "all" Then
14            CCysToBuild = AllCcys
15        Else
16            CCysToBuild = sCompareTwoArrays(AllCcys, _
                  sTokeniseString(RangeFromSheet(shConfig, "CurrenciesToInclude")), _
                  "Common")
17            If sNRows(CCysToBuild) = 1 Then
18                CCysToBuild = Numeraire
19            Else
20                CCysToBuild = sRemoveDuplicates(sArrayStack(Numeraire, sDrop(CCysToBuild, 1)))
21            End If
22        End If

23        CCysToBuild = SortCurrencies(CCysToBuild, Numeraire)
          
          'Save market data to file...
24        MessageLogWrite "Exporting market data for " & sConcatenateStrings(CCysToBuild, ", ") & " to " & Folder
          
25        ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
              MarketWB, FileName, CCysToBuild, Numeraire, , 2)
26        ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
              MarketWB, FileNameH, CCysToBuild, Numeraire, , 3)
27        Exit Sub
ErrHandler:
28        Throw "#ExportMarketData (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

