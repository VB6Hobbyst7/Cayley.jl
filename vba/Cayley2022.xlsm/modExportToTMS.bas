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
          Dim AnchorDate As Date
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
              'Feeding the rates may well change the AnchorDate
15            SolumDotOut = RangeFromSheet(MarketDataWorkbook.Worksheets("Config"), "MarketDataFile")
16            SolumDotOut = sJoinPath(MarketDataWorkbook.Path, SolumDotOut)
17            AnchorDate = AnchorDateFromMarketDataFile(SolumDotOut)
18        Else
19            AnchorDate = RangeFromMarketDataBook("Config", "AnchorDate").Value2
20        End If

21        TargetFolder = RangeFromSheet(shExportToTMS, "WhereToExport", False, True, False, False, False)
22        TargetFolder = sJoinPath(TargetFolder, Format(AnchorDate, "yyyy-mm-dd"))
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
32                    Throw "Cannot find scenario definition file '" & ScenarioList(i, 1) & "'"
33                ElseIf LCase(Right$(ScenarioList(i, 1), 4)) <> ".sdf" Then
34                    Throw "Scenario definition files should have .sdf file extension, but file '" + ScenarioList(i, 1) + "' does not"
35                End If
36            Next i
              Dim Tmp
37            Tmp = sCountDistinctItems(ScenarioList)
38            If Tmp(1, 2) > 1 Then
39                Throw "Scenario file '" + Tmp(1, 1) + "' appears more than once in the list of scenarios to be executed"
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

61        ShowFileInSnakeTail , False 'Does not throw error is SnakeTail not installed.

62        AllBanks = sSortedArray(GetColumnFromLinesBook("CPTY_PARENT", LinesBook))
63        PrepareForCalculation AllBanks(1, 1), True, True, True

64        If DoFeedRates Then
65            FeedRatesFromTextFile
66            BuildModelsInJulia True, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")
67        End If

68        If DoExportMarketData Then
69            ExportMarketData MarketDataFolder, AnchorDate
70        End If

71        If DoExportTrades Then
72            ExportTrades TradesFolder, AnchorDate
73        End If

74        If DoExportCharts Then
75            PrintCharts False, AllBanks, ChartsFolder, ExportJPG, AnchorDate, True
76        End If

77        If DoExportTable Then
78            RunTable DoTradeHeadroom, DoFxHeadroom, True, AllBanks
79            FileNameTable = sJoinPath(TableFolder, _
                  "ResultsByCounterpartyParent_" & Format(AnchorDate, "yyyy-mm-dd") & ".csv")
80            ExportTable FileNameTable
81        End If

82        If DoScenarios Then
83            BuildModelsInJulia False, 1, 1
84            RunManyScenarios ScenarioList, True, MN_CM, gModel_CM, ScenariosFolder
85        End If

86        shExportToTMS.Activate

87        TimeEnd = Now
88        MessageLogWrite "Cayley RunExport started at " + Format(TimeStart, "yyyy-mm-dd hh:mm:ss")
89        MessageLogWrite "Cayley RunExport ended at " + Format(TimeEnd, "yyyy-mm-dd hh:mm:ss")
90        MessageLogWrite "Cayley RunExport took " & Format(TimeEnd - TimeStart, "hh:mm:ss")

91        Exit Sub
ErrHandler:
92        SomethingWentWrong "#RunExport (line " & CStr(Erl) + "): " & Err.Description & "!"
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

3         Prompt = "Export files for Treasury Management System?" + vbLf + vbLf + _
              "The following tasks will be carried out:"
                   
4         If DoFeedRates Then
              Dim SolumDotOut As String
5             AnyTasks = True
6             SolumDotOut = RangeFromSheet(MarketWB.Worksheets("Config"), "MarketDataFile")
7             SolumDotOut = sJoinPath(MarketWB.Path, SolumDotOut)
                   
8             Prompt = Prompt + vbLf + vbLf + _
                  "Rates will be fed from file" + vbLf + _
                  SolumDotOut + vbLf + _
                  "to the market data workbook"
                   
9         End If
            
10        If DoExportTrades Then
11            AnyTasks = True
12            Prompt = Prompt + vbLf + vbLf + _
                  "Three trade data files will be copied to" + vbLf & _
                  TradesFolder
13        End If
              
14        If DoExportMarketData Then
15            AnyTasks = True
16            Prompt = Prompt + vbLf + vbLf + _
                  "Three market data files will be copied to" + vbLf & _
                  MarketDataFolder
17        End If
                  
18        If DoExportTable Then
19            AnyTasks = True
20            Prompt = Prompt + vbLf + vbLf + _
                  "The Table worksheet will be updated to record bank-by-bank trade headroom and fx headroom. Results will be saved to" + vbLf + _
                  TableFolder
21        End If

22        If DoExportCharts Then
23            AnyTasks = True
24            Prompt = Prompt + vbLf + vbLf + _
                  "Charts plotting the current PFE of each bank's trades with Airbus and the bank's lines to Airbus will be saved to" + vbLf + _
                  ChartsFolder
25        End If

26        If DoScenarios Then
27            AnyTasks = True
28            Prompt = Prompt + vbLf + vbLf + _
                  NumberToText(NumScenarios) & " scenario" & IIf(NumScenarios > 1, "s", "") & " will be run with scenario definition files (.sdf) and scenario result files (.srf) saved to:" + vbLf + _
                  ScenariosFolder
29        End If

30        If DoScenarios Or DoExportTable Then
31            AnyTasks = True
32            Prompt = Prompt + vbLf + vbLf + _
                  "The process may take some time. You can monitor progress if SnakeTail is installed." + vbLf + "http://snakenest.com/snaketail/"
33        End If

34        If Not AnyTasks Then
35            Throw "Please select at least one task, from FeedRates through ExportCharts, or at least one scenario.", True
36        End If

37        AreYouSure = MsgBoxPlus(Prompt, vbQuestion + vbOKCancel, m_MsgBoxTitle) = vbOK
38        Exit Function
ErrHandler:
39        Throw "#AreYouSure (line " & CStr(Erl) + "): " & Err.Description & "!"
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
8         Throw "#NumberToText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub ExportTrades(TargetFolder As String, AnchorDate As Date)
          Dim Suffix As String

1         On Error GoTo ErrHandler

2         Suffix = "_" & Format(AnchorDate, "yyyy-mm-dd") & ".csv"

3         ThrowIfError sFileCopy(FileFromConfig("FxTradesCSVFile"), sJoinPath(TargetFolder, "FxTrades" & Suffix))
4         ThrowIfError sFileCopy(FileFromConfig("RatesTradesCSVFile"), sJoinPath(TargetFolder, "RatesTrades" & Suffix))
5         ThrowIfError sFileCopy(FileFromConfig("AmortisationCSVFile"), sJoinPath(TargetFolder, "Amortisation" & Suffix))

6         Exit Sub
ErrHandler:
7         Throw "#ExportTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
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
7         Throw "#SaveETMSToRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetETMSFromRegistry
' Author     : Philip Swannell
' Date       : 22-Mar-2022
' Purpose    : Restore the contents of the worksheet from the Windows registry. Called from the workbook open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub GetETMSFromRegistry()
          Dim i As Long
          Dim RangeNames
          Dim Res
          
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
8                 Res = sParseArrayString(CStr(Res))
9                 Set TargetRange = RangeFromSheet(shExportToTMS, CStr(RangeNames(i, 1)))
10                If TargetRange.Rows.Count = sNRows(Res) And TargetRange.Columns.Count = sNCols(Res) Then
11                    TargetRange.Value = sArrayExcelString(Res)
12                Else
13                    TargetRange.ClearContents
14                    With TargetRange.Resize(sNRows(Res), sNCols(Res))
15                        .Value = Res
16                        shExportToTMS.Names.Add RangeNames(i, 1), .Cells
17                    End With
18                End If
19            End If
20        Next i

21        FormatExportToTMS

22        Exit Sub
ErrHandler:
23        Throw "#GetETMSFromRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
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
14                .ErrorMessage = "Value must be text"
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
40                        .ErrorMessage = "Value must be TRUE or FALSE"
41                        .ShowInput = True
42                        .ShowError = True
43                    End With
44                    If IsEmpty(c.Value) Then c.Value = False
45                End With
46            Next

47        End With

48        With RangeFromSheet(ws, "Scenarios")
49            TopLeftAddress = Replace(.Cells(1, 1).Address, "$", "")
50            AddGreyBorders .Cells, True
51            CayleyFormatAsInput .Cells
52            .Columns(1).HorizontalAlignment = xlHAlignCenter
53            .Columns(2).HorizontalAlignment = xlHAlignLeft
54            .FormatConditions.Delete
55            .FormatConditions.Add Type:=xlExpression, Formula1:="=" & "$" & TopLeftAddress & "<>TRUE"
56            .FormatConditions(.FormatConditions.Count).SetFirstPriority
57            With .FormatConditions(1).Font
58                .ThemeColor = xlThemeColorDark1
59                .TintAndShade = -0.249946592608417
60            End With
61            .FormatConditions(1).StopIfTrue = False

62            With .Columns(1).Validation
63                .Delete
64                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="=ISLOGICAL(" & TopLeftAddress & ")"
65                .IgnoreBlank = True
66                .InCellDropdown = True
67                .InputTitle = "DoThisOne?"
68                .ErrorTitle = "Export to TMS"
69                .InputMessage = "Double-click to change"
70                .ErrorMessage = "Value must be TRUE or FALSE"
71                .ShowInput = True
72                .ShowError = True
73            End With

74            With .Columns(2).Validation
75                .Delete
76                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="=ISTEXT(" & TopLeftAddress & ")"
77                .IgnoreBlank = True
78                .InCellDropdown = True
79                .InputTitle = "ScenarioDefinitionFile"
80                .ErrorTitle = "Export to TMS"
81                .InputMessage = "Double-click to browse"
82                .ErrorMessage = "Value must be text"
83                .ShowInput = True
84                .ShowError = True
85            End With
86        End With
87        Exit Sub
ErrHandler:
88        Throw "#FormatExportToTMS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExportMarketData
' Author     : Philip Swannell
' Date       : 21-Mar-2022
' Purpose    : Exports market data: Three files: 1) The current Solum.out, created by Airbus tech department, its location
'              is taken from the Config sheet of the Market Data Workbook. 2) & 3) the .json files that are parsed by the Julia code
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ExportMarketData(Folder As String, AnchorDate As Date)

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
6             Throw "Cannot find file '" + SolumDotOut + "' Please check the 'MarketDataFile' setting on the 'Config' worksheet of the MarketDataWorkbook"
7         End If

8         ThrowIfError sFileCopy(SolumDotOut, sJoinPath(Folder, "MarketDataFile_" & Format(AnchorDate, "yyyy-mm-dd") & ".out"))

9         FileName = sJoinPath(Folder, "CayleyMarket_" & Format(AnchorDate, "yyyy-mm-dd") & ".json")
10        FileNameH = sJoinPath(Folder, "CayleyMarketHistoricFxVol_" & Format(AnchorDate, "yyyy-mm-dd") & ".json")
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
24        StatusBarWrap "Exporting market data for " & sConcatenateStrings(CCysToBuild, ", ")
          
25        ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
              MarketWB, FileName, CCysToBuild, Numeraire, , 2)
26        ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
              MarketWB, FileNameH, CCysToBuild, Numeraire, , 3)
27        Exit Sub
ErrHandler:
28        Throw "#ExportMarketData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function AnchorDateFromMarketDataFile(SolumDotOut As String)
          Dim FileContents
          Dim i As Long

1         On Error GoTo ErrHandler
2         If Not sFileExists(SolumDotOut) Then Throw "Cannot find file '" + SolumDotOut + "' which is the current setting for 'MarketDataFile' on the Config sheet of the MarketDataWorkbook"

3         FileContents = sCSVRead(SolumDotOut, False, False, , , , , , , , 20)

4         For i = 1 To sNRows(FileContents)
5             If InStr(FileContents(i, 1), "RUNDATE") > 0 Then
6                 AnchorDateFromMarketDataFile = AnchorDateFromRunDate(CStr(FileContents(i, 1)))
7                 Exit Function
8             End If
9         Next i

10        Throw "RUNDATE not found in file '" & SolumDotOut & "' . Should appear as a line in the file of the form 'RUNDATE=YYYYMMDD'"

11        Exit Function
ErrHandler:
12        Throw "#AnchorDateFromMarketDataFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'BAD - this code is copied and pasted from the MarketData workbook. Duplicated code!
Function AnchorDateFromRunDate(RunDateText As String)
          Dim TheDay As String
          Dim TheMonth As String
          Dim TheYear As String
          Dim yyyymmdd As String
          Const ErrString = "RUNDATE not found in file. Should appear as a line in the file of the form 'RUNDATE=YYYYMMDD'"

1         On Error GoTo ErrHandler
2         yyyymmdd = Trim(sStringBetweenStrings(RunDateText, "="))
3         If Len(yyyymmdd) <> 8 Then Throw ErrString
4         If Not sIsRegMatch("^[0-9]*$", yyyymmdd) Then Throw ErrString

5         TheYear = Left(yyyymmdd, 4)
6         TheMonth = Mid(yyyymmdd, 5, 2)
7         TheDay = Right(yyyymmdd, 2)
8         If CLng(TheMonth) > 12 Or CLng(TheMonth) < 1 Then Throw ErrString
9         If CLng(TheDay) > 31 Or CLng(TheDay) < 1 Then Throw ErrString
10        AnchorDateFromRunDate = DateSerial(CInt(TheYear), CInt(TheMonth), CInt(TheDay))

11        Exit Function
ErrHandler:
12        Throw "#AnchorDateFromRunDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


