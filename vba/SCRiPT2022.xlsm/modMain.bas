Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Two methods: ShowMenu that pops up the menu on the Portfolio sheet, the
'             xVADashboard sheet and the the method XVAFrontEndMain, the "Main" of this workbook
'---------------------------------------------------------------------------------------
Option Explicit
Public gTradesAsOfLastPFECalc As Variant
Public gResults As Dictionary
Public Const gDoubleClickPrompt = "<Doubleclick to add trade>"
Public Const gDoValidation = True 'True 'If False then validation of trade data is suppressed so that we can more easily test the Julia code's handling of bad data. Method ReleaseCleanup prevents release of the workbook with this constant set to True

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Author     : Philip Swannell
' Date       : 07-Feb-2018
' Purpose    : Return a writable directory for saving results files to be communicated to Julia. Preference for c:\temp\XVA\ if it's available
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp(Optional Refresh As Boolean = False)
          Static Res As String
1         On Error GoTo ErrHandler
2         If Not Refresh And Res <> "" Then
3             LocalTemp = Res
4             Exit Function
5         End If
6         Res = "c:\temp"
7         If Not sFolderIsWritable(Res) Then
8             Res = sEnvironmentVariable("temp")
9         End If
10        If Right(Res, 1) <> "\" Then Res = Res + "\"
11        Res = Res + "XVA\"
12        ThrowIfError sCreateFolder(Res)
13        LocalTemp = Res
14        Exit Function
ErrHandler:
15        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : XVAFrontEndMain
' Author    : Philip Swannell
' Date      : 04-Nov-2015
' Purpose   : The "Main" of this workbook. Saves Market, Trade and Control data to file,
'             to be picked up by the Julia code. Executes the Julia code (via JuliaExcel)
'             and presents results in the sheets of this workbook.
'---------------------------------------------------------------------------------------
Sub XVAFrontEndMain(ByVal DoPV As Boolean, ByVal DoCVA As Boolean, ByVal DoPFE As Boolean, ByVal DoKVA As Boolean, _
          ByVal PartitionByNetSet As Boolean, ByVal PartitionByTrade As Boolean, _
          ByVal UseCachedModel As Boolean, BuildModelFromDFsAndSurvProbs As Boolean)

          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim MarketWb As Workbook
          Dim NumTrades As Long
          Dim SPH1 As SolumAddin.clsSheetProtectionHandler
          Dim SPH2 As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim Trades As Variant
          Dim TradesForJulia As Variant
          Dim TradesRange As Range
          Dim Ccys As Variant, Credits As Variant, IIs As Variant    'inflation indices
          Dim BanksChosen As Variant
          Dim Numeraire As String
          Dim StatusBarMessage As String

1         On Error GoTo ErrHandler

2         CalculateRequiredMarketData Ccys, IIs
3         If ChooseBanks(True, BanksChosen) = False Then Exit Sub
4         If DoKVA Then Throw "KVA calculation not supported!"

5         If DoCVA And DoPFE Then
              ' ValidateCapitalInputs BanksChosen
6         End If

7         Credits = CreditsRequired()
8         If UseCachedModel Then
9             If sEquals(False, shConfig.Range("UseCachedModel")) Then
10                UseCachedModel = False
11            ElseIf Not (CheckModelHasCurves(Ccys, Credits, IIs)) Then
12                UseCachedModel = False
13            End If
14        End If

15        If Not (UseCachedModel) And Not (DoPFE) And Not (DoCVA) Then
16            StatusBarMessage = "Rebuilding Model..."
17        ElseIf Not (UseCachedModel) And (DoPFE Or DoCVA) Then
18            StatusBarMessage = "Rebuilding Model and Calculating..."
19        Else
20            StatusBarMessage = "Calculating..."
21        End If

22        Set ExSH = CreateExcelStateHandler(xlCalculationManual, , True, StatusBarMessage, , True)
23        Set SUH = CreateScreenUpdateHandler()
24        Set SPH1 = CreateSheetProtectionHandler(shConfig)
25        Set SPH2 = CreateSheetProtectionHandler(shHiddenSheet)

26        Set MarketWb = OpenMarketWorkbook(True, False)
          'Would like to get rid of lines workbook, but not quite possible... Dashboard generation fails... PGS 18/11/20
27        OpenLinesWorkbook True, False

          'Pass configuration data to Julia - tell Julia what to do and how to do it.
28        shConfig.Calculate
29        Numeraire = RangeFromMarketDataBook("Config", "Numeraire")

          Dim BaseFolder As String
          Dim ControlFile As String
          Dim ControlFileFullName As String
          Dim MarketFile As String
          Dim MarketFileFullName As String
          Dim OutputModelFile As String
          Dim ResultsFile As String
          Dim ResultsFileFullName As String
          Dim TradeFile As String
          Dim TradeFileFullName As String
30        BaseFolder = LocalTemp(True) 'has terminating backslash

31        ControlFile = "Control.json" 'Changing these file names? Then change method ShowInputAndOutputFiles!

32        ControlFileFullName = BaseFolder & ControlFile
33        TradeFile = "Trades.csv"
34        TradeFileFullName = BaseFolder & TradeFile
35        ResultsFile = "Results.json"
36        ResultsFileFullName = BaseFolder & ResultsFile

37        MarketFile = "MarketRates.json"
38        MarketFileFullName = BaseFolder & MarketFile
39        OutputModelFile = "Model.jls"
40        SaveControlFile ControlFileFullName, BanksChosen, UseCachedModel, DoPV, DoCVA, DoPFE, PartitionByNetSet, PartitionByTrade, TradeFile, MarketFile, ResultsFile, OutputModelFile
          'Pass trade data to Julia...
41        CalculatePortfolioSheet
42        Set TradesRange = getTradesRange(NumTrades)
43        If NumTrades = 0 Then
44            Trades = Empty
45        Else
46            Trades = TradesRange.Value2
47            If gDoValidation Then
48                ValidateTrades Trades, True
49            End If
50        End If
51        TradesForJulia = PortfolioTradesToJuliaTrades(Trades, True, True)
52        ThrowIfError sFileSaveCSV(TradeFileFullName, TradesForJulia, True, "yyyy-mm-dd")

          'Save market data to file
53        If Not BuildModelFromDFsAndSurvProbs Then
54            If Not (UseCachedModel) Then
55                CheckMarketWorkbook OpenMarketWorkbook(), gProjectName
56                ThrowIfError Application.Run("'" + MarketWb.Name + "'!SaveDataFromMarketWorkbookToFile", MarketWb, MarketFileFullName, Ccys, Numeraire, Credits, False, IIs)
57            End If
58        End If

59        Set gResults = Run_xva_main(ControlFileFullName, ResultsFileFullName, BuildModelFromDFsAndSurvProbs)

          Dim ExSH2 As clsExcelStateHandler
60        Set ExSH2 = CreateExcelStateHandler(, , , "Reading " & BaseFolder & ResultsFile)

62        Set ExSH2 = Nothing
63        If gResults.Exists("Error") Then
              Dim Prompt As String
64            Prompt = "The Julia code failed with error:" + vbLf + vbLf + _
                  gResults("Error") + vbLf + vbLf + _
                  "This was the Julia call stack (root cause at top):" + vbLf + vbLf + _
                  VBA.Join(gResults("Stacktrace"), vbLf)
65            Throw Prompt, True
66        End If

67        If DoPV Then
68            UpdatePortfolioSheet
69        End If

70        If DoPFE Then
71            gTradesAsOfLastPFECalc = TradesRange.Value2
72            UpdateTradeViewerSheet True
73        End If
74        If DoPFE And PartitionByNetSet Then
75            UpdateCounterpartyViewerSheet True
76        End If

77        UpdatexVADashboard False, DoPV, DoCVA, DoKVA, PartitionByNetSet, Numeraire, ConfigRange("OurName").Value

78        If LoggingIsOn Then
79            AppActivate Application.Caption, False
80        End If

81        Exit Sub
ErrHandler:
82        Throw "#XVAFrontEndMain (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveControlFile
' Author     : Philip Swannell
' Date       : 19-Feb-2018
' Purpose    : Saves a json file to be picked up by the Julia code, file is one of "triplet" of files, one for trades, one for market data and this one to control what the Julia code is to do.
' Parameters :
'  ControlFile            : Name of file in Windows format
'  BuildCurvesOnly        : If True then Julia code will build a Hull-White model, but do no valuation
'  CounterpartiesToProcess:
'  UseCachedModel         : if True then Julia code will use a the same HW model (i.e. calibration) as in the previous call - this is "stateful".
'                           if False then Julia code will calibrate a new model from the provided MarketFile
'                           if the name of a file then Julia code will use the contents of that file as the
'                               model. File should be json and will presumambly have been produced by a previous call to the Julia code - See OutputModelFile
'  DoPV                   : Should trade PVs be calculated?
'  DoCVA                  : Should CVAs be calculated
'  DoPFE                  : Should PFEs be calculated?
'  PartitionByNetSet      : If CVA and/or PFEs are to be calculated, should they be calculated for each netset?
'  PartitionByTrade       : If CVA and/or PFEs are to be calculated, should they be calculated for each trade?
'  TradeFile              : The name of the file from which Julia code should read trades
'  MarketFile             : The name of the file from which Julia code should read market data
'  ResultsFile            : The name of the file to which Julia code should write results.
'  OutputModelFile        : The name of the file to which the Julia code writes the model (which is also embedded in ResultsFile)
' -----------------------------------------------------------------------------------------------------------------------
Function SaveControlFile(ControlFile As String, CounterpartiesToProcess As Variant, ByVal UseCachedModel As Boolean, _
          DoPV As Boolean, DoCVA As Boolean, DoPFE As Boolean, PartitionByNetSet As Boolean, PartitionByTrade As Boolean, _
          TradeFile As String, MarketFile As String, ResultsFile As String, OutputModelFile As String)

          Dim DCT As New Dictionary
          Dim JSON As String
          Dim NumSims As Long
          Dim NumSimsCVA As Long
          Dim OnValuationErrors As String
          Dim PFEPercentile As Double
          Dim SavePaths As Boolean
          Dim TimeGap As Double

1         On Error GoTo ErrHandler

2         TimeGap = ConfigRange("TimeGap")
3         PFEPercentile = (1 - ConfigRange("PFEPercentile"))
4         NumSims = ConfigRange("NumSims")
5         NumSimsCVA = ConfigRange("NumSimsCVA")
6         SavePaths = ConfigRange("SavePaths")
7         OnValuationErrors = ConfigRange("OnValuationErrors")

8         DCT.Add "CounterpartiesToProcess", To1D(CounterpartiesToProcess)
9         DCT.Add "DoPV", DoPV
10        DCT.Add "DoCVA", DoCVA
11        DCT.Add "DoPFE", DoPFE
12        DCT.Add "PartitionByNetSet", PartitionByNetSet
13        DCT.Add "PartitionByTrade", PartitionByTrade
14        DCT.Add "TimeGap", TimeGap
15        DCT.Add "PFEPercentile", PFEPercentile
16        DCT.Add "NumSims", NumSims
17        DCT.Add "NumSimsCVA", NumSimsCVA
18        DCT.Add "SelfPartyName", gSELF
19        DCT.Add "WhatIfPartyName", gWHATIF
20        DCT.Add "SavePaths", SavePaths
21        DCT.Add "OnValuationErrors", OnValuationErrors
22        DCT.Add "TradeFile", MorphSlashes(TradeFile, UseLinux())
23        If UseCachedModel Then
24            DCT.Add "InputModelFile", MorphSlashes(OutputModelFile, UseLinux())
25        Else
              'We always generate market data to build from rates + cds spreads. Julia code can morph the control file _
               and market file appropriately to test xva_main's ability to accept market data that contains discount factors and survival proabilities.
26            DCT.Add "BuildCurvesFromRates", True
27            DCT.Add "BuildSurvProbsFromSpreads", True
28            DCT.Add "MarketFile", MorphSlashes(MarketFile, UseLinux())
29        End If
30        DCT.Add "ResultsFile", MorphSlashes(ResultsFile, UseLinux())
31        DCT.Add "OutputModelFile", MorphSlashes(OutputModelFile, UseLinux())
32        DCT.Add "AllowScalarsInResults", False 'Only required by Julia code but Julia code will ignore

33        JSON = ConvertToJson(DCT, 3, AS_RowByRow)

34        ThrowIfError sFileSave(ControlFile, JSON, "")

35        Exit Function
ErrHandler:
36        Throw "#SaveControlFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShowMenu
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Menu from either Portfolio sheet or from xvaDashboard
'             Mode can take values FromMenuButton,FromRightClick and AddOneTrade
'---------------------------------------------------------------------------------------
Sub ShowMenu(Optional Mode As String = "FromMenuButton")

1         On Error GoTo ErrHandler
          
          Dim Activate As Boolean
          Dim BuildModelFromDFsAndSurvProbs As Boolean
          Dim Chosen As Variant
          Dim EnableFlags
          Dim i As Long
          Dim IncludeFileOptions As Boolean
          Dim IncludeNonTradingOptions As Boolean
          Dim IncludeTradingOptions As Boolean
          Dim Message As String
          Dim NumTrades As Long
          Dim NumTradesToAdd As Long
          Dim PartitionByTrade As Boolean
          Dim t1 As Double
          Dim TheChoices
          Dim TheFids

          'WHEN ADDING NEW TRADE TYPE, REMEMBER TO INCREASE CONSTANT NumTradeTypes
          Const chAddInterestRateSwap = "Interest Rate S&wap"
          Const FidAddInterestRateSwap = 137
          Const chAddCrossCurrencySwap = "&Cross Currency Swap"
          Const FidAddCrossCurrencySwap = 137
          Const chAddFixedCashflows = "F&ixed Cashflows"
          Const FidAddFixedCashflows = 137
          Const chAddFxForward = "&Fx Forward"
          Const FidAddFxForward = 137
          Const chAddFxForwardStrip = "Fx Forwa&rd Strip"
          Const FidAddFxForwardStrip = 137
          Const chAddFxOption = "F&x Option"
          Const FidAddFxOption = 137
          Const chAddFxOptionStrip = "Fx Optio&n Strip"
          Const FidAddFxOptionStrip = 137
          Const chAddCapFloor = "Ca&p Floor"
          Const FidAddCapFloor = 137
          Const chAddSwaption = "Swap&tion"
          Const FidAddSwaption = 137
          Const chAddInflationZCSwap = "Inflation &ZC Swap"
          Const FidAddInflationZCSwap = 137
          Const chAddInflationYoYSwap = "Inflation &YoY Swap"
          Const FidAddInflationYoYSwap = 137
          Const chDelete = "--&Delete selected trades..."
          Const FidDelete = 358
          Const chSolve = "For PV = &0"
          Const FidSolve = 70
          Const chSolve2 = "For PV = &Target..."
          Const FidSolve2 = 5827
          Const chCashflows = "Cas&hflows..."
          Const FidCashflows = 333
          Const chOpenTrades = "&Open Trades File"
          Const chOpenBackups = "&Restore trades from backups..."
          Const FidOpenBackups = 128
2         Dim chSave As String: chSave = "&Save Trades"
          Const FidSave = 3
          Const chSaveAs = "Save Trades &As..."
          Const FidSaveAs = 3
          Const chClose = "&Close Trades File"
          Const FidClose = 923
          Const chRepairTradeIDS = "--R&epair duplicate Trade IDs"
          Const FidRepairTradeIDs = 548
          Dim chBuildModel As String
          Dim FidBuildModel As Long
          Const chChooseCPs = "Choose &Banks..."
          Const FidChooseCPs = 221
          Const chPV = "Calculate P&Vs                                              (Shift F9)"
          Const FidPV = 12604    '50
          Const chPVCVA = "Calculate &xVAs"
          Const FidPVCVA = 156
          Dim chPVCVAPFE As String
          Const FidPVCVAPFE = 156

          Const chOpenMarket = "Open &Market Data workbook...            (Shift to activate)"
          Const FidOpenMarket = 23
          Const chActivateMarket = "Activate &Market Data workbook"
          Const FidActivateMarket = 142
          Const chOpenLines = "Open &Lines workbook...                        (Shift to activate)"
          Const FidOpenLines = 23
          Const chActivateLines = "Activate &Lines workbook"
          Const FidActivateLines = 142
          Const chCreateSystemImage = "Create Julia System Image..."
          Const FidCreateSystemImage = 11208
          Const chLaunchJuliaWithoutSystemImage = "Launch Julia without System Image"
          Const FidLaunchJuliaWithoutSystemImage = 16330
          
          Const chShowInsAndOuts = "Show &Input and Output files"
          Const FidShowInsAndOuts = 5
          Const chShowTradesForR = "Sho&w Trades formatted for Julia"
          Const FidShowTradesForR = 12351
          Const chTestTradeConversion = "Test trade con&version..."
          Const FidTestTradeConversion = 37
          Const chCompress = "&Compress trades"
          Const FidCompress = 3987
          Const chUncompress = "&Uncompress trades"
          Const FidUncompress = 3986
          Const chCreateTestSet = "Create &test set..."
          Const FidCreateTestSet = 0
          
          Dim CompressEnabled As Boolean
          Dim TradesRange As Range
          Dim UncompressEnabled As Boolean

3         Application.Cursor = xlDefault     'because we are at top level, and debugging sessions may have left the cursor in the wrong state
4         Application.EnableCancelKey = xlDisabled    ' This looks dangerous, (and it is dangerous if we introduce _
                                                        an infinite loop in the VBA code) but the Julia code is not interruptable, _
                                                        and that's where most of the time is spent.
5         Application.StatusBar = False   'Only necessary following unhandled errors earlier - should only see on development PC
6         gBlockCalculateEvent = False    ' This is safe since we know this method must be at the top of the call stack!
7         gBlockChangeEvent = False    'ditto
8         Set TradesRange = getTradesRange(NumTrades)
9         If Mode = "FromMenuButton" Then
10            IncludeFileOptions = True
11            If ActiveSheet.Name = shPortfolio.Name Then
12                IncludeTradingOptions = True
13                IncludeNonTradingOptions = True
14            Else
15                IncludeTradingOptions = False
16                IncludeNonTradingOptions = True
17            End If
18        ElseIf Mode = "FromRightClick" Then
19            IncludeFileOptions = True
20            IncludeTradingOptions = True
21            IncludeNonTradingOptions = False
22        ElseIf Mode = "AddOneTrade" Then
23            IncludeFileOptions = False
24            IncludeTradingOptions = True
25            IncludeNonTradingOptions = False
26            Application.StatusBar = "Shift to add more than one trade"
27        Else
28            Throw "Mode not recognised"
29        End If

30        PartitionByTrade = ConfigRange("PartitionByTrade").Value
31        BuildModelFromDFsAndSurvProbs = ConfigRange("BuildModelFromDFsAndSurvProbs").Value

32        If IncludeNonTradingOptions Then
33            If ModelExists() Then
34                chBuildModel = "Rebuild Hull-&White Model"
35                FidBuildModel = 37
36            Else
37                chBuildModel = "Build Hull-&White Model"
38                FidBuildModel = 5828
39            End If
40            If PartitionByTrade Then
41                chPVCVAPFE = "Calculate xVAs and &PFE (inc. by trade)"
42            Else
43                chPVCVAPFE = "Calculate xVAs and &PFE"
44            End If
45            TheChoices = sArrayStack(chBuildModel, chChooseCPs, "--" & chPV, chPVCVA, chPVCVAPFE)
46            TheFids = sArrayStack(FidBuildModel, FidChooseCPs, FidPV, FidPVCVA, FidPVCVAPFE)
47            EnableFlags = sArrayStack(True, True, sReshape(NumTrades > 0, 3, 1))
48        Else
49            TheChoices = CreateMissing(): TheFids = CreateMissing(): EnableFlags = CreateMissing()
50        End If

51        If IncludeTradingOptions Then

              Dim TradeChoices
52            TradeChoices = sArrayStack(chAddCapFloor, chAddCrossCurrencySwap, chAddFxForward, chAddFxOption, chAddInterestRateSwap, chAddInflationYoYSwap, _
                  chAddInflationZCSwap, chAddSwaption, sArrayRange("T&rade Strips", sArrayStack(chAddFixedCashflows, chAddFxForwardStrip, chAddFxOptionStrip)))
53            TheFids = sArrayStack(TheFids, FidAddCapFloor, FidAddCrossCurrencySwap, FidAddFxForward, FidAddFxOption, FidAddInterestRateSwap, FidAddInflationYoYSwap, _
                  FidAddInflationZCSwap, FidAddSwaption, FidAddFixedCashflows, FidAddFxForwardStrip, FidAddFxOptionStrip)

54            If IncludeNonTradingOptions Then    'Make a sub-menu
55                TheChoices = sArrayStack(TheChoices, sArrayRange("--&New Trade(s)", TradeChoices))
56            Else    ' put them on the top-level menu
57                TheChoices = sArrayStack(TheChoices, TradeChoices)
58            End If
59            EnableFlags = sArrayStack(EnableFlags, sReshape(True, sNRows(TradeChoices), 1))
60        End If

61        If IncludeFileOptions Then
              Dim MRUChoices
              Dim MRUEnabled
              Dim MRUFids
              Dim MRUFiles
              Dim SaveAllowed As Boolean
62            GetMRUList gProjectName & "TradeFiles", MRUFiles, MRUChoices, MRUFids, MRUEnabled, False
63            AmendMRU MRUFiles, MRUChoices
64            SaveAllowed = (NumTrades > 0) And LCase(Right(RangeFromSheet(shPortfolio, "TradesFileName").Value, 4)) = ".stf" And _
                  (sSplitPath(CStr(RangeFromSheet(shPortfolio, "TradesFileName").Value), False) & "\" <> BackUpDirectory)
65            If SaveAllowed Then chSave = chSave & "   (as " & sSplitPath(RangeFromSheet(shPortfolio, "TradesFileName")) & ")"

66            If IsEmpty(MRUFiles) Then   'no files used recently
67                TheChoices = sArrayStack(TheChoices, sArrayRange("--" & chOpenTrades, sArrayStack("--Browse...", "--" & chOpenBackups)))
68                TheFids = sArrayStack(TheFids, 23, FidOpenBackups)
69                EnableFlags = sArrayStack(EnableFlags, True, True)
70            Else
71                TheChoices = sArrayStack(TheChoices, sArrayRange("--" & chOpenTrades, sArrayStack(MRUChoices, "--" & chOpenBackups)))
72                TheFids = sArrayStack(TheFids, MRUFids, FidOpenBackups)
73                EnableFlags = sArrayStack(EnableFlags, MRUEnabled, True)
74            End If
75            TheChoices = sArrayStack(TheChoices, chSave, chSaveAs, chClose)
76            TheFids = sArrayStack(TheFids, FidSave, FidSaveAs, FidClose)
77            EnableFlags = sArrayStack(EnableFlags, SaveAllowed, NumTrades > 0, NumTrades > 0)
78            TheChoices = sArrayStack(TheChoices, chDelete, sArrayRange("Sol&ve selected trades", sArrayStack(chSolve, chSolve2)), chCashflows, chRepairTradeIDS)
79            TheFids = sArrayStack(TheFids, FidDelete, FidSolve, FidSolve2, FidCashflows, FidRepairTradeIDs)
80            EnableFlags = sArrayStack(EnableFlags, NumSelectedTrades() > 0, sReshape(sEquals(True, AreSelectedTradesSolvable()), 2, 1), (NumSelectedTrades() > 0 And ModelExists()), TradeIDsNeedRepairing())
81        End If

82        If IncludeNonTradingOptions Then
              Dim LinesBookOpen As Boolean
              Dim MarketBookOpen As Boolean
83            MarketBookOpen = IsInCollection(Application.Workbooks, sSplitPath(FileFromConfig("MarketDataWorkbook")))
84            LinesBookOpen = IsInCollection(Application.Workbooks, sSplitPath(FileFromConfig("LinesWorkbook")))
85            TheChoices = sArrayStack(TheChoices, "--" & IIf(MarketBookOpen, chActivateMarket, chOpenMarket), IIf(LinesBookOpen, chActivateLines, chOpenLines))
86            TheFids = sArrayStack(TheFids, IIf(MarketBookOpen, FidActivateMarket, FidOpenMarket), IIf(LinesBookOpen, FidActivateLines, FidOpenLines))
87            EnableFlags = sArrayStack(EnableFlags, True, True)
              Dim AdvancedChoices
              Dim AdvancedEnabled
88            AdvancedChoices = sArrayRange("--Developer &Tools", _
                  sArrayStack(chCreateSystemImage, chLaunchJuliaWithoutSystemImage, chShowInsAndOuts, "--" & chShowTradesForR, chTestTradeConversion, "--" & chCreateTestSet))
89            AdvancedEnabled = sReshape(True, sNRows(AdvancedChoices), 1)
90            TheChoices = sArrayStack(TheChoices, AdvancedChoices, sArraySquare("Compress&ion", chCompress, "", chUncompress))
91            TheFids = sArrayStack(TheFids, FidCreateSystemImage, FidLaunchJuliaWithoutSystemImage, FidShowInsAndOuts, FidShowTradesForR, FidTestTradeConversion, FidCreateTestSet, FidCompress, FidUncompress)
92            EnableCompressMenuItems CompressEnabled, UncompressEnabled, TradesRange, NumTrades
93            EnableFlags = sArrayStack(EnableFlags, AdvancedEnabled, CompressEnabled, UncompressEnabled)
94        End If

95        Chosen = ShowCommandBarPopup(TheChoices, TheFids, EnableFlags)

96        If Not IsEmpty(MRUChoices) Then
97            For i = 1 To sNRows(MRUFiles)
98                If Chosen = Unembellish(CStr(MRUChoices(i, 1))) Then
99                    If OpenTradesFile(CStr(MRUFiles(i, 1)), True) Then
100                       AddFileToMRU gProjectName & "TradeFiles", CStr(MRUFiles(i, 1))
101                       BackUpTrades    'opening a trades file is a "big change" so back up immediately
102                   End If
103                   Exit Sub
104               End If
105           Next i
106       End If

          'back up trades if they haven't been backed up in the last three minutes
107       If NumTrades > 0 Then
108           If (Now - g_LastTradeBackUpTime) > 3 / 24 / 60 Then
109               If shPortfolio.g_LastChangeTime > g_LastTradeBackUpTime Then
110                   BackUpTrades
111               End If
112           End If
113       End If

114       If Mode = "AddOneTrade" Then NumTradesToAdd = IIf(IsShiftKeyDown(), 0, 1)

115       t1 = sElapsedTime()
116       Select Case Chosen
              Case "#Cancel!"
117               Exit Sub
118           Case Unembellish(chAddInterestRateSwap)
119               AddTrades "InterestRateSwap", NumTradesToAdd
120           Case Unembellish(chAddCrossCurrencySwap)
121               AddTrades "CrossCurrencySwap", NumTradesToAdd
122           Case Unembellish(chAddFixedCashflows)
123               AddTrades "FixedCashflows", NumTradesToAdd
124           Case Unembellish(chAddCapFloor)
125               AddTrades "CapFloor", NumTradesToAdd
126           Case Unembellish(chAddSwaption)
127               AddTrades "Swaption", NumTradesToAdd
128           Case Unembellish(chAddFxForward)
129               AddTrades "FxForward", NumTradesToAdd
130           Case Unembellish(chAddFxOption)
131               AddTrades "FxOption", NumTradesToAdd
132           Case Unembellish(chAddFxForwardStrip)
133               AddTrades "FxForwardStrip", NumTradesToAdd
134           Case Unembellish(chAddFxOptionStrip)
135               AddTrades "FxOptionStrip", NumTradesToAdd
136           Case Unembellish(chAddInflationYoYSwap)
137               AddTrades "InflationYoYSwap", NumTradesToAdd
138           Case Unembellish(chAddInflationZCSwap)
139               AddTrades "InflationZCSwap", NumTradesToAdd
140           Case Unembellish(chDelete)
141               DeleteSelectedTrades
142           Case Unembellish(chSolve)
143               SolveSelectedTrades
144           Case Unembellish(chSolve2)
145               SolveSelectedTradesForTargetPV
146           Case Unembellish(chRepairTradeIDS)
147               RepairTradeIDs
148           Case Unembellish(chOpenTrades), "Browse..."
149               If OpenTradesFile(, True) Then
150                   BackUpTrades    'opening a trades file is a "big change" so back up immediately
151               End If
152           Case Unembellish(chSaveAs)
153               SaveTradesFile
154           Case Unembellish(chSave)
155               SaveTradesFile RangeFromSheet(shPortfolio, "TradesFileName")
156           Case Unembellish(chPV)
157               XVAFrontEndMain True, False, False, False, False, False, True, BuildModelFromDFsAndSurvProbs
158           Case Unembellish(chBuildModel)
159               XVAFrontEndMain False, False, False, False, False, False, False, BuildModelFromDFsAndSurvProbs
160           Case Unembellish(chChooseCPs)
161               ChooseBanks False
162           Case Unembellish(chPVCVA)
163               PartitionByTrade = False
164               XVAFrontEndMain True, True, False, False, True, PartitionByTrade, True, BuildModelFromDFsAndSurvProbs
165               shxVADashboard.Activate
166           Case Unembellish(chPVCVAPFE)
167               XVAFrontEndMain True, True, True, False, True, PartitionByTrade, True, BuildModelFromDFsAndSurvProbs
168               shxVADashboard.Activate
169           Case Unembellish(chOpenMarket)
170               Activate = IIf(MarketBookOpen, True, IsShiftKeyDown())
171               OpenMarketWorkbook False, Activate
172           Case Unembellish(chActivateMarket)
173               OpenMarketWorkbook False, True
174           Case Unembellish(chOpenLines)
175               OpenLinesWorkbook False, IsShiftKeyDown()
176           Case Unembellish(chActivateLines)
177               OpenLinesWorkbook False, True
178           Case Unembellish(chShowTradesForR)
179               ShowTradesForJulia
180           Case Unembellish(chCashflows)
181               ViewTradeCashflows Application.Intersect(ActiveCell.EntireRow, TradesRange)
182           Case Unembellish(chTestTradeConversion)
183               TestTradeConversion
184           Case Unembellish(chCompress)
185               CompressTradesOnPortfolioSheet
186           Case Unembellish(chUncompress)
187               UncompressTradesOnPortfolioSheet
188           Case Unembellish(chOpenBackups)
189               OpenBackUps
190           Case "Clear Recent File List"
191               RemoveFileFromMRU gProjectName & "TradeFiles", "All"
192           Case Unembellish(chShowInsAndOuts)
193               ShowInputAndOutputFiles
194           Case Unembellish(chClose)
195               CloseTrades
196           Case Unembellish(chCreateTestSet)
197               CreateTestSet
198           Case Unembellish(chCreateSystemImage)
199               JuliaCreateSystemImage True, UseLinux()
200           Case Unembellish(chLaunchJuliaWithoutSystemImage)
201               LaunchJuliaWithoutSystemImage
202           Case Else
203               MsgBoxPlus "Unrecognised choice:" + Chosen, vbCritical, MsgBoxTitle()
204       End Select

205       If Mode = "AddOneTrade" Then
206           Application.StatusBar = False
207       ElseIf Chosen <> Chosen <> Unembellish(chSolve) And Chosen <> Unembellish(chCashflows) Then    'these ones do their own temporary message
              'Where chosen contains a sequence of spaces, assume that's because the text appearing in the menu includes mention of an accelerator key e.g. "Calculate P&Vs                        (Shift F9)"
208           Message = "Time to " + sStringBetweenStrings(Chosen, , "     ") + ": " + Format(sElapsedTime() - t1, "0.00") + " seconds"
209           TemporaryMessage Message, 4, ""
210       End If

211       Exit Sub
ErrHandler:
212       SomethingWentWrong "#ShowMenu (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, MsgBoxTitle()
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowInputAndOutputFiles
' Author     : Philip Swannell
' Date       : 16-Mar-2018
' Purpose    : Mainly as a utility for Metasite, allows them to see the input and output files
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowInputAndOutputFiles()
          Dim ControlFile As String
          Dim MarketFileDiscountFactors As String
          Dim MarketFileRates As String
          Dim OutputModelFile As String
          Dim ResultsFile As String
          Dim TradeFile As String
1         On Error GoTo ErrHandler
2         ControlFile = LocalTemp(True) & "Control.json"
3         TradeFile = LocalTemp() & "Trades.csv"
4         ResultsFile = LocalTemp() & "Results.json"
5         MarketFileRates = LocalTemp() & "MarketRates.json"
6         MarketFileDiscountFactors = LocalTemp() & "MarketDiscountFactors.json"
7         OutputModelFile = LocalTemp() & "Model.jls"
8         ShowFileInTextEditor ControlFile
9         ShowFileInTextEditor MarketFileRates
10        ShowFileInTextEditor MarketFileDiscountFactors
11        ShowFileInTextEditor TradeFile
12        ShowFileInTextEditor ResultsFile
13        Exit Sub
ErrHandler:
14        Throw "#ShowInputAndOutputFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' ----------------------------------------------------------------
' Procedure Name: AmendMRU
' Purpose: Amend the "MRUChoices" for recently accessed trade files to include a brief description of the file's contents
' Parameter MRUFiles (): single column of file names, as generated by a call to GetMRUList
' Parameter MRUChoices ():Single column of text strings to appear in command bar, as returned by GetMRUList, this method overwrites this byRef variable
' Author: Philip Swannell
' Date: 05-Dec-2017
' ----------------------------------------------------------------
Private Sub AmendMRU(ByVal MRUFiles, ByRef MRUChoices)
          Dim i As Long
          Dim LeftPart As String
          Dim NumSpaces As Long
          Dim RightPart As String
          Dim Summary As Variant

1         On Error GoTo ErrHandler
2         If Not IsEmpty(MRUFiles) Then
3             For i = 1 To sNRows(MRUFiles)
4                 Summary = TradeFileInfo(CStr(MRUFiles(i, 1)), "TradesSummary")
5                 If VarType(Summary) = vbString Then
6                     If Not sIsErrorString(Summary) Then
7                         RightPart = AbbreviateTradeSummary(CStr(Summary))
8                         If i <= 9 Then
9                             LeftPart = "&" & CStr(i) + " " + CStr(MRUFiles(i, 1))
10                        ElseIf i = 10 Then
11                            LeftPart = "1&0" + " " + CStr(MRUFiles(i, 1))
12                        Else
13                            LeftPart = CStr(MRUFiles(i, 1))
14                        End If
                          'Aim of the code below is to Right-justify the RightPart in the command bar, it doesn't quite work, suggesting that the font used in command bars is not Segoe UI, 9 point?
15                        NumSpaces = CLng((390 - sStringWidth(LeftPart & RightPart, "Segoe UI", 9)(1, 1)) / 2.25 - 0.5)
16                        If NumSpaces < 1 Then NumSpaces = 1
17                        MRUChoices(i, 1) = AbbreviateForCommandBar(LeftPart & String(NumSpaces, " ") & RightPart)
18                    End If
                      ' Debug.Print sStringWidth(MRUChoices(i, 1), "Segoe UI", 9)(1, 1), MRUChoices(i, 1)
19                End If
20            Next i
21        End If
22        Exit Sub
ErrHandler:
23        Throw "#AmendMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnableCompressMenuItems
' Author    : Philip Swannell
' Date      : 10-Jan-2017
' Purpose   : Determine whether the menu options to compress and to uncompress trades
'             should be enabled.
'---------------------------------------------------------------------------------------
Private Sub EnableCompressMenuItems(ByRef CompressEnabled As Boolean, ByRef UncompressEnabled As Boolean, TradesRange As Range, NumTrades As Long)
          Dim FindRes As Variant
          Dim VFs As Variant
          Static VFsToSearchFor

1         On Error GoTo ErrHandler
2         CompressEnabled = False
3         UncompressEnabled = False

4         If NumTrades = 0 Then Exit Sub
5         VFs = TradesRange.Columns(gCN_TradeType).Value

6         If IsEmpty(VFsToSearchFor) Then
7             VFsToSearchFor = sArrayStack("FxForward", "FxOption", "FxForwardStrip", "FxOptionStrip")
8         End If

9         FindRes = sArrayIsNumber(sMatch(VFsToSearchFor, VFs))

10        CompressEnabled = sColumnOr(FindRes)(1, 1)
11        UncompressEnabled = FindRes(3, 1) Or FindRes(4, 1)
12        Exit Sub
ErrHandler:
13        Throw "#EnableCompressMenuItems (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


