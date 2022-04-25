Attribute VB_Name = "modScenario"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : WhoHasLinesFromButton
' Author    : Philip Swannell
' Date      : 11-Jun-2015
' Purpose   : Runs WhoHasLines taking inputs from the PFE sheet. Running from the button
'             is mainly done for debugging purposes. We don't expose the WhoHasLines sheet
'             to the user.
' -----------------------------------------------------------------------------------------------------------------------
Sub WhoHasLinesFromButton()
1         On Error GoTo ErrHandler
          Dim AllocationsByYear As String
          Dim FirstRow As Long
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim HedgeHorizon As Long
          Dim IncludeAssetClasses As String
          Dim LastRow As Long
          Dim LinesScaleFactor As Double
          Dim NumMCPaths As Long
          Dim NumObservations As Long
          Dim PortfolioAgeing As Double
          Dim Prompt As String
          Dim PromptArray
          Dim TradesScaleFactor As Double
          Const ThrowErrors = False

2         IncludeAssetClasses = RangeFromSheet(shCreditUsage, "IncludeAssetClasses", False, True, False, False, False)
3         NumMCPaths = RangeFromSheet(shCreditUsage, "NumMCPaths", True, False, False, False, False).Value2
4         NumObservations = RangeFromSheet(shCreditUsage, "NumObservations", True, False, False, False, False).Value2
5         FxShock = RangeFromSheet(shCreditUsage, "FxShock", True, False, False, False, False).Value2
6         FxVolShock = RangeFromSheet(shCreditUsage, "FxVolShock", True, False, False, False, False).Value2
7         PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False).Value2
8         FirstRow = RangeFromSheet(shWhoHasLines, "First_row_to_run", True, False, False, False, False).Value
9         LastRow = RangeFromSheet(shWhoHasLines, "Last_row_to_run", True, False, False, False, False).Value
10        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value
11        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False).Value
12        AllocationsByYear = RangeFromSheet(shWhoHasLines, "AllocationsByYear", False, True, False, False, False).Value
13        HedgeHorizon = GetHedgeHorizon()

14        If FirstRow < 1 Then FirstRow = 1
15        If LastRow < 1 Or LastRow > (RangeFromSheet(shWhoHasLines, "TheTable").Rows.Count - 1) Then
16            LastRow = RangeFromSheet(shWhoHasLines, "TheTable").Rows.Count - 1
17        End If

18        PromptArray = sArrayStack("First row to run", FirstRow, _
              "Last row to run", LastRow, _
              "", "")
19        If PortfolioAgeing <> 0 Then
20            PromptArray = sArrayStack(PromptArray, _
                  "PortfolioAgeing", Format(PortfolioAgeing, "0.000"))

21        End If
22        PromptArray = sArrayStack(PromptArray, _
              "FxShock", FxShock, _
              "FxVolShock", FxVolShock, _
              "IncludeAssetClasses", IncludeAssetClasses, _
              "NumMCPaths", Format(NumMCPaths, "###,###"), _
              "NumObservations", Format(NumObservations, "###,###"), _
              "IncludeFutureTrades", True, _
              "FilterBy2", CStr(RangeFromSheet(shCreditUsage, "FilterBy2").Value), _
              "Filter2Value", CStr(RangeFromSheet(shCreditUsage, "Filter2Value").Value), _
              "CurrenciesToInclude", RangeFromSheet(shConfig, "CurrenciesToInclude"), _
              "HedgeHorizon", HedgeHorizon, _
              "AllocationsByYear", AllocationsByYear)

23        If TradesScaleFactor <> 1 Or LinesScaleFactor <> 1 Then
24            PromptArray = sArrayStack(PromptArray, _
                  "", "", _
                  "Morphing:", "", _
                  "TradesScaleFactor", TradesScaleFactor, _
                  "LinesScaleFactor", LinesScaleFactor)
25        End If

26        PromptArray = sReshape(PromptArray, sNRows(PromptArray) / 2, 2)
27        PromptArray = CleanUpPromptArray(PromptArray, True)

28        Prompt = "Run ""Who has Lines"" with the following inputs:" & vbLf & _
              sConcatenateStrings(sJustifyArrayOfStrings(PromptArray, "Calibri", 11, vbTab), vbLf)

29        If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, "Who has lines?") <> vbOK Then Exit Sub

30        OpenOtherBooks
31        JuliaLaunchForCayley
32        BuildModelsInJulia False, FxShock, FxVolShock

33        WhoHasLines NumMCPaths, NumObservations, PortfolioAgeing, "", FxShock, FxVolShock, AllocationsByYear, _
              FirstRow, LastRow, ThrowErrors, HedgeHorizon
34        Exit Sub
ErrHandler:
35        SomethingWentWrong "#WhoHasLinesFromButton (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : WhoHasLines
' Author    : Philip Swannell
' Date      : 10-Jun-2015
' Purpose   : Loops through all the banks, recalculating the PFE sheet in order to figure
'             out which banks will be able to execute further trades. Results are pasted
'             to the sheet WhoHasLines, and picked up from there by subsequent calls to
'             AllocateTradesToBanks.
' -----------------------------------------------------------------------------------------------------------------------
Sub WhoHasLines(NumMCPaths As Long, NumObservations As Long, _
          PortfolioAgeing As Double, StatusBarPrefix As String, FxShock As Double, _
          FxVolShock As Double, AllocationsByYear As String, Optional ByVal FirstRow As Long, _
          Optional ByVal LastRow As Long, Optional ThrowErrors As Boolean, Optional ByVal HedgeHorizon As Long)

          Dim Allocations As Variant
          Dim CopyOfErr As String
          Dim D As Dictionary
          Dim i As Long
          Dim Message As String
          Dim RangeToAutoFit As Range
          Dim Res As Variant
          Dim SPH As Object
          Dim SPH2 As Object
          Dim SUH As Object
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double
          Dim TheBank As String
          Dim TheRange As Range

1         On Error GoTo ErrHandler

2         t1 = sElapsedTime()
3         Set SPH = CreateSheetProtectionHandler(shWhoHasLines)
4         Set SPH2 = CreateSheetProtectionHandler(shCreditUsage)
5         Set SUH = CreateScreenUpdateHandler()
6         Allocations = ParseAllocation(AllocationsByYear)

7         If HedgeHorizon = 0 Then HedgeHorizon = GetHedgeHorizon()

8         AmendWhoHasLinesTable HedgeHorizon

9         With RangeFromSheet(shWhoHasLines, "TheTable")
10            Set TheRange = .offset(1).Resize(.Rows.Count - 1)
11        End With

12        If FirstRow < 1 Then FirstRow = 1
13        If LastRow < 1 Or LastRow > TheRange.Rows.Count Then LastRow = TheRange.Rows.Count

14        With RangeFromSheet(shWhoHasLines, "Time_of_run")
15            .Value = Now()
16            .NumberFormat = "dd-mmm-yyyy hh:mm:ss"
17        End With
          
18        RangeFromSheet(shWhoHasLines, "FxShock").Value = FxShock
19        RangeFromSheet(shWhoHasLines, "FxVolShock").Value = FxVolShock
20        RangeFromSheet(shWhoHasLines, "PortfolioAgeing").Value = PortfolioAgeing
21        RangeFromSheet(shWhoHasLines, "NumObservations").Value = NumObservations
22        RangeFromSheet(shWhoHasLines, "NumMcPaths").Value = NumMCPaths
23        RangeFromSheet(shWhoHasLines, "TradesScaleFactor").Value = RangeFromSheet(shCreditUsage, "TradesScaleFactor").Value
24        RangeFromSheet(shWhoHasLines, "LinesScaleFactor").Value = RangeFromSheet(shCreditUsage, "LinesScaleFactor").Value
25        RangeFromSheet(shWhoHasLines, "IncludeAssetClasses").Value = RangeFromSheet(shCreditUsage, "IncludeAssetClasses").Value
26        RangeFromSheet(shWhoHasLines, "CurrenciesToInclude").Value = RangeFromSheet(shConfig, "CurrenciesToInclude").Value
27        RangeFromSheet(shWhoHasLines, "HedgeHorizon").Value = HedgeHorizon
28        RangeFromSheet(shWhoHasLines, "LastRunAllocationsByYear").Value = AllocationsByYear
29        RangeFromSheet(shWhoHasLines, "LastRunfirst_row_to_run").Value = FirstRow
30        RangeFromSheet(shWhoHasLines, "LastRunlast_row_to_run").Value = LastRow

          'Prepare the sheets
31        RangeFromSheet(shCreditUsage, "NumMCPaths").Value = NumMCPaths
32        RangeFromSheet(shCreditUsage, "NumObservations") = NumObservations
33        RangeFromSheet(shCreditUsage, "IncludeExtraTrades") = False
34        RangeFromSheet(shCreditUsage, "FilterBy1").Value = "Counterparty Parent"
35        RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value = 0
36        RangeFromSheet(shCreditUsage, "PortfolioAgeing").Value = PortfolioAgeing
37        RangeFromSheet(shCreditUsage, "IncludeFutureTrades").Value = True
38        RangeFromSheet(shCreditUsage, "FxShock").Value = FxShock
39        RangeFromSheet(shCreditUsage, "FxVolShock").Value = FxVolShock

          'When using the HW model, there is a significant speed up in headroom solving _
           (cacheing of the "cube" of trade values and unit hedge trade values) that works best when successive banks _
           being solved for have the same BaseCurrency and also use the same model (historic vols vs market implied _
           vols). So we re-order the banks.
          'After moving to Julia (Feb 2022), this may no longer be necessary, since we have more sophisticated _
           memoizing, but reordering doesn't hurt.

          Dim ArrayToSort
          Dim colBanks
          Dim colBCs
          Dim colIndices
          Dim N As Long
          Dim RowToProcess As Long
40        colBanks = shWhoHasLines.Range("TheTable").Columns(1).offset(FirstRow).Resize(LastRow - FirstRow + 1).Value
41        Force2DArrayR colBanks
42        colBCs = sReshape(0, LastRow - FirstRow + 1, 1)
43        colIndices = colBCs
44        For i = 1 To LastRow - FirstRow + 1
45            colIndices(i, 1) = FirstRow + i - 1
46        Next i
47        colBCs = sArrayTranspose(LookupCounterpartyInfo(colBanks, "Base Currency"))
          Dim ColIsH
          Dim ColIsNB
48        ColIsNB = sArrayEquals("NOTIONAL BASED", sArrayTranspose(LookupCounterpartyInfo(colBanks, "METHODOLOGY")))
49        ColIsH = sArrayEquals("HISTORICAL", sArrayTranspose(LookupCounterpartyInfo(colBanks, "Volatility Input")))

50        ArrayToSort = sArrayRange(ColIsNB, ColIsH, colBCs, colIndices, colBanks)
51        ArrayToSort = sSortedArray(ArrayToSort, 1, 2, 3, False, True, True)
52        N = sNRows(ArrayToSort)

53        With TheRange
54            .offset(FirstRow - 1, 1).Resize(LastRow - FirstRow + 1, .Columns.Count - 1).ClearContents
55            For i = 1 To LastRow - FirstRow + 1
56                RowToProcess = ArrayToSort(i, 4)
57                TheBank = .Cells(RowToProcess, 1)
58                RangeFromSheet(shCreditUsage, "Filter1Value") = TheBank
59                Message = StatusBarPrefix & "Bank " & CStr(i) & " of " & CStr(LastRow - FirstRow + 1) & " " & TheBank & _
                      " " & ArrayToSort(i, 3) & _
                      IIf(ArrayToSort(i, 1), " Notional-based", IIf(ArrayToSort(i, 2), " Monte Carlo Historic Vol", " Monte Carlo Implied Vol")) & _
                      " " & ArrayToSort(i, 5)

60                MessageLogWrite Message
61                Set D = New Dictionary
62                t3 = sElapsedTime()
63                Res = RunCreditUsageSheet("Solve345", ThrowErrors, False, False, D, Allocations)
64                t4 = sElapsedTime()

65                If D.Exists("TradeSolveResult") Then
66                    If D("TradeSolveResult") <> "OK" Then
67                        Res = D("TradeSolveResult")
68                    End If
69                End If
                                       
70                With .Cells(RowToProcess, 2).Resize(1, HedgeHorizon)
71                    If sIsErrorString(Res) Then
72                        .Value = Res
73                    ElseIf D.Exists("TradeHeadroom345") Then
74                        .Value = sArrayTranspose(D("TradeHeadroom345"))
75                    End If
76                End With

77                .Cells(RowToProcess, HedgeHorizon + 3) = t4 - t3
78                If D.Exists("PVUSD") Then
79                    .Cells(RowToProcess, HedgeHorizon + 2) = D("PVUSD")
80                Else
81                    .Cells(RowToProcess, HedgeHorizon + 2) = "Cannot find key 'PVUSD' in results dictionary"
82                End If

83            Next i

84        End With
85        TheBank = ""

86        Set RangeToAutoFit = Application.Union(TheRange, _
              Range(shWhoHasLines.Range("Time_of_run").offset(0, -1), _
              shWhoHasLines.Range("LastRunLast_row_to_run")))
87        RangeToAutoFit.Columns.AutoFit
88        t2 = sElapsedTime()
89        Exit Sub
ErrHandler:
90        CopyOfErr = "#WhoHasLines (line " & CStr(Erl) & "): " & Err.Description
91        If TheBank <> "" Then
92            CopyOfErr = CopyOfErr & " (bank being processed = " & TheBank & ")!"
93        Else
94            CopyOfErr = CopyOfErr & "!"
95        End If
96        Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AmendWhoHasLinesTable
' Author     : Philip Swannell
' Date       : 14-Feb-2022
' Purpose    : Fixes up the columns of range TheTable on sheet WhohHasLines in case the HedgeHorizon is different from
'              the last time it was run
' Parameters :
'  HedgeHorizon:
' -----------------------------------------------------------------------------------------------------------------------
Sub AmendWhoHasLinesTable(HedgeHorizon As Long)

          Dim i As Long
          Dim NumColsToAdd As Long
          Dim R As Range
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set R = RangeFromSheet(shWhoHasLines, "TheTable")
3         NumColsToAdd = HedgeHorizon + 3 - R.Columns.Count

4         If NumColsToAdd = 0 Then
5             Exit Sub
6         ElseIf NumColsToAdd > 0 Then
7             Set SPH = CreateSheetProtectionHandler(shWhoHasLines)
8             On Error Resume Next
9             RemoveSortButtonsInRange R.Cells(0, 1).EntireRow
10            On Error GoTo ErrHandler
11            R.Columns(R.Columns.Count - 1).Resize(, NumColsToAdd).Insert Shift:=xlToRight, _
                  CopyOrigin:=xlFormatFromLeftOrAbove
12            For i = 1 To HedgeHorizon
13                R.Cells(1, 1 + i) = CStr(i) & "Y"
14            Next i
15            Set R = RangeFromSheet(shWhoHasLines, "TheTable")
16            On Error Resume Next
17            AddSortButtons R.Rows(0), 1
18            On Error GoTo ErrHandler
19        ElseIf NumColsToAdd < 0 Then
20            Set SPH = CreateSheetProtectionHandler(shWhoHasLines)
21            On Error Resume Next
22            RemoveSortButtonsInRange R.Cells(0, 1).EntireRow
23            On Error GoTo ErrHandler
24            R.Columns(R.Columns.Count - 1 + NumColsToAdd).Resize(, -NumColsToAdd).Delete Shift:=xlToLeft
25            Set R = RangeFromSheet(shWhoHasLines, "TheTable")
26            On Error Resume Next
27            AddSortButtons R.Rows(0), 1
28            On Error GoTo ErrHandler
29        End If
30        Exit Sub
ErrHandler:
31        Throw "#AmendWhoHasLinesTable (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub Test_ExecuteTrades()
          Dim AvEDSTraded As Double
          Dim CallRatio As Double
          Dim CallStrikeOffset As Double
          Dim CCY1
          Dim CCY2
          Dim Ccy2Notionals
          Dim Counterparties
          Dim DealDate As Long
          Dim ForwardsRatio As Double
          Dim FxShock
          Dim MaturityLabels
          Dim PortfolioAgeing As Double
          Dim PutRatio As Double
          Dim PutStrikeOffset As Double
          Dim UnshockedModelBareBones As Dictionary
          Dim UnshockedModelName As String

1         On Error GoTo ErrHandler
2         If gModel_CM Is Nothing Then
3             OpenOtherBooks
4             BuildModelsInJulia True, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")
5         End If

6         DealDate = DateSerial(2021, 11, 27)
7         Counterparties = sTokeniseString("UBSW_CH_ZRH,LLCM_GB_LON,BOTK_JP_TYO,CHAS_US_NYC,LBBW_DE_STR,BSCH_ES_MAD,HELA_DE_FRA,GENO_DE_FRA,HSBC_GB_LON,ROYC_CA_TOR,NOSC_CA_TOR,GSGI_US_NYC,MHCB_JP_TYO,COBA_DE_FRA,NBAD_AE_AUH,BKCH_CN_BJS,ANZB_AU_MEL,UOVB_SG_SIN,UBSW_CH_ZRH,LLCM_GB_LON,BOTK_JP_TYO,CHAS_US_NYC,LBBW_DE_STR,BSCH_ES_MAD,HELA_DE_FRA,GENO_DE_FRA,HSBC_GB_LON,ROYC_CA_TOR,NOSC_CA_TOR,GSGI_US_NYC,MHCB_JP_TYO,COBA_DE_FRA,NBAD_AE_AUH,BKCH_CN_BJS,ANZB_AU_MEL,UOVB_SG_SIN,UBSW_CH_ZRH,LLCM_GB_LON,BOTK_JP_TYO,CHAS_US_NYC,LBBW_DE_STR,BSCH_ES_MAD,HELA_DE_FRA,GENO_DE_FRA,HSBC_GB_LON,ROYC_CA_TOR,NOSC_CA_TOR,GSGI_US_NYC,MHCB_JP_TYO,COBA_DE_FRA,NBAD_AE_AUH,BKCH_CN_BJS,ANZB_AU_MEL,UOVB_SG_SIN")
8         CCY1 = sReshape("EUR", 54, 1)
9         CCY2 = sReshape("USD", 54, 1)
10        Ccy2Notionals = sArrayStack(-50000000, -50000000, -50000000, -50000000, -50000000, -50000000, -50000000, _
              -39615329.9325407, -50000000, -50000000, -50000000, -50000000, -50000000, -23275561.9244161, -50000000, _
              -50000000, -50000000, -20442441.4763765, -50000000, -50000000, -50000000, -50000000, -50000000, -50000000, _
              -50000000, -39615329.9325407, -50000000, -50000000, -50000000, -50000000, -50000000, -23275561.9244161, _
              -50000000, -50000000, -50000000, -20442441.4763765, -50000000, -50000000, -50000000, -50000000, -50000000, _
              -50000000, -50000000, -39615329.9325407, -50000000, -50000000, -50000000, -50000000, -50000000, _
              -23275561.9244161, -50000000, -50000000, -50000000, -20442441.4763765)
11        MaturityLabels = sTokeniseString("3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,3Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,4Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y,5Y")
12        FxShock = 0.965393794749403
13        AvEDSTraded = 1.169082009
14        ForwardsRatio = 0.6
15        PutRatio = -0.2
16        CallRatio = 0.2
17        PutStrikeOffset = -0.1
18        CallStrikeOffset = 0.1
19        UnshockedModelName = "cayleymodel"
20        PortfolioAgeing = 31 / 365
21        Set UnshockedModelBareBones = gModel_CM

22        ExecuteTrades DealDate, Counterparties, CCY1, CCY2, Ccy2Notionals, MaturityLabels, FxShock, AvEDSTraded, ForwardsRatio, PutRatio, CallRatio, PutStrikeOffset, CallStrikeOffset, UnshockedModelName, PortfolioAgeing, UnshockedModelBareBones

23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#Test_ExecuteTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExecuteTrades
' Author    : Philip Swannell
' Date      : 09-Jun-2015
' Purpose   : Executes a number of ATM Forward Fx trades, saving the trade data into the
'             sheet FutureTrades.
'             All variant arguments are column arrays of the same height.
'             Arguments ForwardsRatio thru PutStikeOffset allow for implementation of
'            "Options Strategies" - e.g. do 80% Forward and 20% Collars
' -----------------------------------------------------------------------------------------------------------------------
Function ExecuteTrades(DealDate As Long, Counterparties As Variant, ByVal CCY1 As Variant, _
          ByVal CCY2 As Variant, ByVal Ccy2Notionals As Variant, _
          MaturityLabels As Variant, ByVal FxShock, AvEDSTraded As Double, _
          ForwardsRatio As Double, PutRatio As Double, CallRatio As Double, _
          PutStrikeOffset As Double, CallStrikeOffset As Double, UnshockedModelName As String, _
          PortfolioAgeing As Double, UnshockedModelBareBones As Dictionary)

          Dim AdjMaturities
          Dim AnchorDate As Long
          Dim BuySell_Out
          Dim CCY1_Out As Variant
          Dim Ccy1Notionals_Out As Variant
          Dim CCY2_Out As Variant
          Dim Ccy2Notionals_Out
          Dim Counterparties_Out
          Dim FirstTradeID As Double
          Dim Forwards
          Dim Headers As Variant
          Dim Maturities
          Dim Maturities_Out
          Dim NumTrades As Long
          Dim NumTradesToWrite
          Dim ProductType_Out
          Dim PutCall_Out
          Dim SPH As Object
          Dim Strikes_Out
          Dim TargetCell As Range
          Dim TradeIDs
          Dim UnshockedSpot

1         On Error GoTo ErrHandler

2         NumTrades = sNRows(Ccy2Notionals)
3         NumTradesToWrite = (IIf(ForwardsRatio <> 0, 1, 0) + IIf(CallRatio <> 0, 1, 0) + IIf(PutRatio <> 0, 1, 0)) * NumTrades

4         CCY1_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), CCY1), _
              IIf(PutRatio = 0, createmissing(), CCY1), _
              IIf(CallRatio = 0, createmissing(), CCY1))

5         CCY2_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), CCY2), _
              IIf(PutRatio = 0, createmissing(), CCY2), _
              IIf(CallRatio = 0, createmissing(), CCY2))

6         ProductType_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), sReshape("FXForward", NumTrades, 1)), _
              IIf(PutRatio = 0, createmissing(), sReshape("FXOption", NumTrades, 1)), _
              IIf(CallRatio = 0, createmissing(), sReshape("FXOption", NumTrades, 1)))
              
7         PutCall_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), sReshape(createmissing(), NumTrades, 1)), _
              IIf(PutRatio = 0, createmissing(), sArrayConcatenate("PUT ", CCY1)), _
              IIf(CallRatio = 0, createmissing(), sArrayConcatenate("CALL ", CCY1)))

8         BuySell_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), sReshape(IIf(ForwardsRatio < 0, "Sell", "Buy"), NumTrades, 1)), _
              IIf(PutRatio = 0, createmissing(), sReshape(IIf(PutRatio < 0, "Sell", "Buy"), NumTrades, 1)), _
              IIf(CallRatio = 0, createmissing(), sReshape(IIf(CallRatio < 0, "Sell", "Buy"), NumTrades, 1)))

9         Ccy2Notionals_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), sArrayMultiply(Ccy2Notionals, ForwardsRatio)), _
              IIf(PutRatio = 0, createmissing(), sArrayMultiply(Ccy2Notionals, PutRatio)), _
              IIf(CallRatio = 0, createmissing(), sArrayMultiply(Ccy2Notionals, CallRatio)))

10        Counterparties_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), Counterparties), _
              IIf(PutRatio = 0, createmissing(), Counterparties), _
              IIf(CallRatio = 0, createmissing(), Counterparties))

          'Prevent coding error.
11        If GetItem(UnshockedModelBareBones, "fxshock") <> 1 Then Throw "Model must be passed in unshocked state"
12        If GetItem(UnshockedModelBareBones, "fxvolshock") <> 1 Then Throw "Model must be passed in unshocked state"

13        AnchorDate = GetItem(UnshockedModelBareBones, "AnchorDate")
14        UnshockedSpot = GetItem(UnshockedModelBareBones, "EURUSD")

15        AdjMaturities = TenureStringsToDates(AnchorDate, MaturityLabels)
16        Maturities = sArrayAdd(AdjMaturities, PortfolioAgeing * 365)

17        Forwards = EURUSDForwardRates(AdjMaturities, UnshockedModelName)
18        If AvEDSTraded <> 0 Then
19            FxShock = AvEDSTraded / UnshockedSpot
20        End If

          'Forwards are calculated in Unshocked market, so we need to apply shocks here
21        Forwards = sArrayMultiply(Forwards, FxShock)

22        Headers = sArrayTranspose(RangeFromSheet(shFutureTrades, "Headers").Value)
23        Set SPH = CreateSheetProtectionHandler(shFutureTrades)

24        Set TargetCell = RangeFromSheet(shFutureTrades, "Headers").Cells(1, 1).End(xlDown).End(xlDown).End(xlUp).offset(1)
25        FirstTradeID = TargetCell.Row - shFutureTrades.Range("Headers").Row
26        TradeIDs = sGrid(FirstTradeID, FirstTradeID + NumTradesToWrite - 1, CLng(NumTradesToWrite))
27        TradeIDs = sArrayConcatenate("FT", TradeIDs)

28        Maturities_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), Maturities), _
              IIf(PutRatio = 0, createmissing(), Maturities), _
              IIf(CallRatio = 0, createmissing(), Maturities))

29        Strikes_Out = Interleave(IIf(ForwardsRatio = 0, createmissing(), Forwards), _
              IIf(PutRatio = 0, createmissing(), sArrayAdd(Forwards, PutStrikeOffset)), _
              IIf(CallRatio = 0, createmissing(), sArrayAdd(Forwards, CallStrikeOffset)))

30        Ccy1Notionals_Out = sArrayDivide(Ccy2Notionals_Out, Strikes_Out)
31        Ccy1Notionals_Out = sArrayMultiply(Ccy1Notionals_Out, -1)

32        With TargetCell.offset(0, sMatch("Trade Id", Headers) - 1).Resize(NumTradesToWrite)
33            .Value = TradeIDs
34        End With
35        With TargetCell.offset(0, sMatch("Prim Amt", Headers) - 1).Resize(NumTradesToWrite)
36            .Value = Ccy1Notionals_Out
37            .NumberFormat = "#,##0;[Red]-#,##0"
38        End With
39        With TargetCell.offset(0, sMatch("Sec Amt", Headers) - 1).Resize(NumTradesToWrite)
40            .Value = Ccy2Notionals_Out
41            .NumberFormat = "#,##0;[Red]-#,##0"
42        End With
43        With TargetCell.offset(0, sMatch("Prim Cur", Headers) - 1).Resize(NumTradesToWrite)
44            .Value = CCY1_Out
45        End With
46        With TargetCell.offset(0, sMatch("Sec Cur", Headers) - 1).Resize(NumTradesToWrite)
47            .Value = CCY2_Out
48        End With
49        With TargetCell.offset(0, sMatch("Product Type", Headers) - 1).Resize(NumTradesToWrite)
50            .Value = ProductType_Out
51        End With
52        With TargetCell.offset(0, sMatch("Settle Date", Headers) - 1).Resize(NumTradesToWrite)
53            .Value = DealDate
54            .NumberFormat = "dd-mmm-yyyy"
55        End With
56        With TargetCell.offset(0, sMatch("PAWE", Headers) - 1).Resize(NumTradesToWrite)
57            .Value = PortfolioAgeing
58        End With
59        With TargetCell.offset(0, sMatch("Maturity Date", Headers) - 1).Resize(NumTradesToWrite)
60            .Value = Maturities_Out
61            .NumberFormat = "dd-mmm-yyyy"
62        End With
63        With TargetCell.offset(0, sMatch("Counterparty Parent", Headers) - 1).Resize(NumTradesToWrite)
64            .Value = Counterparties_Out
65        End With
66        With TargetCell.offset(0, sMatch("Buy Sell", Headers) - 1).Resize(NumTradesToWrite)
67            .Value = BuySell_Out
68        End With
69        With TargetCell.offset(0, sMatch("Put Call", Headers) - 1).Resize(NumTradesToWrite)
70            .Value = PutCall_Out
71        End With
          
72        shFutureTrades.Names.Add "TheTrades", TargetCell.CurrentRegion.offset(1).Resize(TargetCell.CurrentRegion.Rows.Count - 1)

73        Exit Function
ErrHandler:
74        Throw "#ExecuteTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Interleave
' Author    : Philip Swannell
' Date      : 13-Jul-2015
' Purpose   : Array1, Array2 and Array3 should be single column arrays.
'   If Array1 = {1;2;3}
'      Array2 = {10;20;30}
'      Array3 = {100;200;3000}
'Then return is  the "interleaving" i.e. {1;10;100;2;20;200;3;30;300}
' -----------------------------------------------------------------------------------------------------------------------
Private Function Interleave(Optional Array1, Optional Array2, Optional Array3)
          Dim Res
1         On Error GoTo ErrHandler

2         Res = sArrayRange(Array1, Array2, Array3)
3         Interleave = sReshape(Res, sNRows(Res) * sNCols(Res), 1)

4         Exit Function
ErrHandler:
5         Throw "#Interleave (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearFutureTrades
' Author    : Philip Swannell
' Date      : 09-Jun-2015
' Purpose   : Removes all trades held on the FutureTrades sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearFutureTrades()
          Dim Res
          Dim SPH As Object
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shFutureTrades)
3         Res = shFutureTrades.UsedRange.Rows.Count        'Reset Used Range
4         shFutureTrades.UsedRange.offset(RangeFromSheet(shFutureTrades, "Headers").Row).EntireRow.Delete
5         If IsInCollection(shFutureTrades.Names, "TheTrades") Then
6             shFutureTrades.Names("TheTrades").Delete
7         End If
8         Res = shFutureTrades.UsedRange.Rows.Count        'Reset Used Range
9         Exit Sub
ErrHandler:
10        Throw "#ClearFutureTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TenureStringsToDates
' Author    : Philip Swannell
' Date      : 09-Jun-2015
' Purpose   : Interpret strings such as "3Y" as specifying a date from a start date.
'             Very basic date handling but could become more sophisticated if necessary,
'             e.g. rollday conventions (such as "Mod Foll"), offsets etc.
'            TenureStrings should be column array
' -----------------------------------------------------------------------------------------------------------------------
Function TenureStringsToDates(FromDate As Long, TenureStrings)
          Dim i As Long
          Dim N As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Force2DArray TenureStrings
3         N = sNRows(TenureStrings)
4         Result = sReshape(0, N, 1)
5         For i = 1 To N
6             If Right(TenureStrings(i, 1), 1) = "Y" Then
7                 Result(i, 1) = FromDate + 365 * CLng(Left(TenureStrings(i, 1), Len(TenureStrings(i, 1)) - 1))
8             Else
9                 Throw "Unrecognised TenureString"
10            End If
11        Next i
12        TenureStringsToDates = Result

13        Exit Function
ErrHandler:
14        Throw "#TenureStringsToDates (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestAllocateTradesToBanks
' Author    : Philip Swannell
' Date      : 29-Sep-2016
' Purpose   : Trivial test harness
' -----------------------------------------------------------------------------------------------------------------------
Sub TestAllocateTradesToBanks()
          Dim CCY1
          Dim CCY2
          Dim Ccy2Notionals
          Dim Counterparties
          Dim MaturityLabels
          Dim NumTradesDone As Long
          Dim TotalAvailableAfterAllocation
          Dim TotalsDone
1         On Error GoTo ErrHandler
2         AllocateTradesToBanks 10000000000#, 10000000000#, 13000000000#, 13000000000#, 13000000000#, _
              10000000000#, 10000000000#, 0, 0, 0, _
              8, Counterparties, CCY1, CCY2, Ccy2Notionals, MaturityLabels, TotalsDone, _
              TotalAvailableAfterAllocation, NumTradesDone
3         g sArrayRange(Counterparties, CCY1, CCY2, Ccy2Notionals, MaturityLabels, TotalsDone)
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestAllocateTradesToBanks (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AllocateTradesToBanks
' Author    : Philip Swannell
' Date      : 09-Jun-2015
' Purpose   : Assumes WhoHasLines has been run. This is the method that we want to tweak to
'             select a "strategy" for intelligently allocating trades. But start with Naive algorithm
' -----------------------------------------------------------------------------------------------------------------------
Function AllocateTradesToBanks(ByVal Amount1Y, ByVal Amount2Y, ByVal Amount3Y, ByVal Amount4Y, ByVal Amount5Y, _
          ByVal Amount6Y, ByVal Amount7Y, ByVal Amount8Y, ByVal Amount9Y, ByVal Amount10Y, HedgeHorizon As Long, _
          ByRef Counterparties As Variant, ByRef CCY1 As Variant, ByRef CCY2 As Variant, _
          ByRef Ccy2Notionals As Variant, ByRef MaturityLabels As Variant, ByRef TotalsDone As Variant, _
          ByRef TotalAvailableAfterAllocation, ByRef NumTradesDone As Long)

          Dim Amount As Double
          Dim i As Long
          Dim j As Long
          Dim Label As String
          Dim NR As Long
          Dim SomeoneHasLines As Boolean
          Dim ThisClip As Double
          Dim WhoHasLinesArray
          Const ClipSize = 50000000
          Dim STK_Ccy2Nots As Object
          Dim STK_CP As Object
          Dim STK_ML As Object

          Dim ColNo_LastY As Long

1         On Error GoTo ErrHandler

2         ColNo_LastY = HedgeHorizon + 1

3         Set STK_CP = CreateStacker()
4         Set STK_Ccy2Nots = CreateStacker()
5         Set STK_ML = CreateStacker()        'ML stands for Maturity Label

6         With RangeFromSheet(shWhoHasLines, "TheTable")
7             WhoHasLinesArray = .offset(1).Resize(.Rows.Count - 1).Value
8         End With
9         WhoHasLinesArray = sSortedArray(WhoHasLinesArray, ColNo_LastY, , , False)

          'Flip strings etc to Zero
10        NR = sNRows(WhoHasLinesArray)
11        For i = 1 To NR
12            For j = ColNo_LastY - 2 To ColNo_LastY
13                If Not IsNumber(WhoHasLinesArray(i, j)) Then
14                    WhoHasLinesArray(i, j) = 0
15                End If
16            Next j
17        Next i

18        TotalsDone = sReshape(0, 1, HedgeHorizon)
19        TotalAvailableAfterAllocation = sReshape(0, 1, HedgeHorizon)
20        NumTradesDone = 0

21        For j = 2 To HedgeHorizon + 1
22            i = 1
23            SomeoneHasLines = False
24            Amount = Choose(j - 1, Amount1Y, Amount2Y, Amount3Y, Amount4Y, Amount5Y, _
                  Amount6Y, Amount7Y, Amount8Y, Amount9Y, Amount10Y)
25            Label = CStr(j - 1) & "Y"
26            Do While Amount > 0
27                ThisClip = SafeMin(ClipSize, Amount)
28                ThisClip = SafeMin(ThisClip, WhoHasLinesArray(i, j))

29                If WhoHasLinesArray(i, j) > 0 Then
30                    STK_CP.StackData WhoHasLinesArray(i, 1)
31                    STK_Ccy2Nots.StackData ThisClip
32                    STK_ML.StackData Label
33                    Amount = Amount - ThisClip
34                    SomeoneHasLines = True
35                    WhoHasLinesArray(i, j) = WhoHasLinesArray(i, j) - ThisClip
36                    TotalsDone(1, j - 1) = TotalsDone(1, j - 1) + ThisClip
37                    NumTradesDone = NumTradesDone + 1
38                End If
39                If i = NR Then        ' check that on this pass through the banks we managed to find at least one bank who has lines...
40                    If Not SomeoneHasLines Then
                          'Throw "Lines with all banks are full!"
41                        Exit Do
42                    End If
43                    SomeoneHasLines = False        ' Flip back so that we can check on the subsequent pass through the banks
44                End If
45                i = i Mod NR + 1
46            Loop
47        Next j

48        TotalAvailableAfterAllocation = sColumnSum(sSubArray(WhoHasLinesArray, 1, 2, , HedgeHorizon))

49        If NumTradesDone > 0 Then
50            Counterparties = STK_CP.Report
51            CCY1 = sReshape("EUR", sNRows(Counterparties), 1)
52            CCY2 = sReshape("USD", sNRows(Counterparties), 1)
53            Ccy2Notionals = sArrayMultiply(-1, STK_Ccy2Nots.Report)        'we measure the lines in USD but trades we execute have negative USD notional
54            MaturityLabels = STK_ML.Report
55        Else
56            Counterparties = "Dummy Bank"
57            CCY1 = "EUR"
58            CCY2 = "USD"
59            Ccy2Notionals = 0
60            MaturityLabels = "1Y"
61        End If
          'We've amended the trades on the FutureTrades sheet so make sure that method GetTradesInJuliaFormat cleans out its cached return
62        FlushStatics

63        Exit Function
ErrHandler:
64        Stop
65        Resume
66        Throw "#AllocateTradesToBanks (line " & CStr(Erl) & "): " & Err.Description & "!"
   End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunScenario
' Author    : Philip Swannell
' Date      : 10-Jun-2015
' Purpose   : Combines methods WhoHasLines, AllocateTradesToBanks and ExecuteTrades to
'             march time forward executing trades monthly, according to the ScenarioDefinition.
' -----------------------------------------------------------------------------------------------------------------------
Function RunScenario(SilentMode As Boolean, StatusBarPrePrefix As String, ModelName As String, _
          SaveDefinition As Boolean, SaveResults As Boolean, ModelBareBones As Dictionary)

          Dim AllocationByYear As String
          Dim AnnualReplenishment As Double
          Dim AvEDSTraded As Double
          Dim CallRatio As Variant
          Dim CallStrikeOffset As Variant
          Dim CopyOfErr As String
          Dim dCallRatio As Double
          Dim dCallStrikeOffset As Double
          Dim dForwardsRatio As Double
          Dim dPutRatio As Double
          Dim dPutStrikeOffset As Double
          Dim ForwardsRatio As Variant
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim HighFxSpeed As Double
          Dim i As Long
          Dim InCleanUp As Boolean
          Dim IncludeAssetClasses As String
          Dim j As Long
          Dim LinesScaleFactor As Double
          Dim LowFxSpeed As Double
          Dim ModelType As String
          Dim NumMCPaths As Long
          Dim NumObservations As Long
          Dim NumTradesDone As Long
          Dim PortfolioAgeing As Double
          Dim PutRatio As Variant
          Dim PutStrikeOffset As Variant
          Dim ScenarioDefinition
          Dim ScenarioDescription As String
          Dim ScenDefnHeaders
          Dim SpeedGridBaseVol As Double
          Dim SpeedGridWidth As Double
          Dim StatusBarPrefix As String
          Dim StrategySwitchPoints As Variant
          Dim SUH As Object
          Dim TradeHeadroomInBillions As Double
          Dim TradesScaleFactor As Double
          Dim UseSpeedGrid As Boolean
          Dim VaryGridWidth As Boolean

          Dim AllocationsArray As Variant
          Dim AnchorDate As Long
          Dim BaseSpot As Double
          Dim BaseVol As Double
          Dim ComputerName As String
          Dim CurrenciesToInclude As String
          Dim Filter2Value
          Dim FilterBy2 As String
          Dim HedgeHorizon As Long
          Dim HistoryEnd As Long
          Dim HistoryStart As Long
          Dim ScenarioResults
          Dim ScenarioResultsHeaders
          Dim SheetToActivate As Worksheet
          Dim ShocksDerivedFrom As String
          Dim ThisDealDate As Long
          Dim TimeEnd As Variant
          Dim TimeStart As Variant

          Dim N As Long
          Dim Prompt As String
          Dim PromptArray
          Dim SPH As Object
          Dim SPH2 As Object
          Dim SPH3 As Object
          Dim SPH4 As clsSheetProtectionHandler
          Dim TotalAvailableAfterAllocation
          Dim TotalsDone

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler()
3         Set SPH = CreateSheetProtectionHandler(shScenarioDefinition)
4         Set SPH2 = CreateSheetProtectionHandler(shCreditUsage)
5         Set SPH3 = CreateSheetProtectionHandler(shFutureTrades)
6         Set SPH4 = CreateSheetProtectionHandler(shScenarioResults)
7         Set SheetToActivate = ActiveSheet

8         RefreshScenarioDefinition True        'Reads current spot levels so that shocks are correctly calibrated, also does validation
9         If SaveDefinition Then
              Dim FileName As String
10            FileName = RangeFromSheet(shConfig, "ScenarioResultsDirectory")
11            If Right(FileName, 1) <> "\" Then FileName = FileName & "\"
              Dim Res

12            Res = sCreateFolder(FileName)

13            If sIsErrorString(Res) Then
14                Throw "Error when attempting to create a directory for scenario results: " & Res & vbLf & _
                      "Please check that 'ScenarioResultsDirectory' on the Config sheet is set to a location to which you have write access."
15            End If

16            If Not sFolderIsWritable(FileName) Then
17                Throw "You do not have write access to the scenario results folder." & vbLf & _
                      "Please check that 'ScenarioResultsDirectory' on the Config sheet is set to a location to which you have write access."
18            End If

19            FileName = FileName + ScenarioDescriptionToFileName(RangeFromSheet(shScenarioDefinition, "ScenarioDescription"), "sdf")
20            SaveScenarioDefinitionFile FileName, ""
21        End If

          'Read data from various sheets - in the order that it appears on the ScenarioResults sheet
22        ShocksDerivedFrom = RangeFromSheet(shScenarioDefinition, "ShocksDerivedFrom", False, True, False, False, False).Value
23        If LCase(ShocksDerivedFrom) = "history" Then
24            HistoryStart = RangeFromSheet(shScenarioDefinition, "HistoryStart", True, False, False, False, False).Value
25            HistoryEnd = RangeFromSheet(shScenarioDefinition, "HistoryEnd", True, False, False, False, False).Value
26        Else
27            HistoryStart = Empty: HistoryEnd = Empty
28        End If
29        BaseSpot = RangeFromSheet(shScenarioDefinition, "BaseSpot", True, False, False, False, False).Value
30        BaseVol = RangeFromSheet(shScenarioDefinition, "BaseVol", True, False, False, False, False).Value
31        ForwardsRatio = RangeFromSheet(shScenarioDefinition, "ForwardsRatio", True, True, False, False, False).Value
32        PutRatio = RangeFromSheet(shScenarioDefinition, "PutRatio", True, True, False, False, False).Value
33        CallRatio = RangeFromSheet(shScenarioDefinition, "CallRatio", True, True, False, False, False).Value
34        PutStrikeOffset = RangeFromSheet(shScenarioDefinition, "PutStrikeOffset", True, True, False, False, False).Value
35        CallStrikeOffset = RangeFromSheet(shScenarioDefinition, "CallStrikeOffset", True, True, False, False, False).Value
36        StrategySwitchPoints = RangeFromSheet(shScenarioDefinition, "StrategySwitchPoints", True, True, False, True, False).Value
37        AllocationByYear = RangeFromSheet(shScenarioDefinition, "AllocationByYear", False, True, False, False, False).Value

          'Line below appends extra zeros if necessary
38        AllocationByYear = sConcatenateStrings(ParseAllocation(AllocationByYear, False, False), ":")

39        UseSpeedGrid = RangeFromSheet(shScenarioDefinition, "UseSpeedGrid", False, False, True, False, False).Value
40        SpeedGridWidth = RangeFromSheet(shScenarioDefinition, "SpeedGridWidth", True, False, False, False, False).Value
41        HighFxSpeed = RangeFromSheet(shScenarioDefinition, "HighFxSpeed", True, False, False, False, False).Value
42        LowFxSpeed = RangeFromSheet(shScenarioDefinition, "LowFxSpeed", True, False, False, False, False).Value
43        VaryGridWidth = RangeFromSheet(shScenarioDefinition, "VaryGridWidth", False, False, True, False, False).Value
44        SpeedGridBaseVol = RangeFromSheet(shScenarioDefinition, "SpeedGridBaseVol", True, False, False, False, False).Value
45        AnnualReplenishment = RangeFromSheet(shScenarioDefinition, "AnnualReplenishment", True, False, False, False, False).Value
46        ScenarioDescription = RangeFromSheet(shScenarioDefinition, "ScenarioDescription", False, True, False, False, False).Value
47        ModelType = RangeFromSheet(shCreditUsage, "ModelType", False, True, False, False, False).Value
48        NumMCPaths = RangeFromSheet(shCreditUsage, "NumMCPaths", True, False, False, False, False).Value
49        NumObservations = RangeFromSheet(shCreditUsage, "NumObservations", True, False, False, False, False).Value
50        FilterBy2 = RangeFromSheet(shCreditUsage, "FilterBy2", False, True, False, False, False).Value
51        Filter2Value = RangeFromSheet(shCreditUsage, "Filter2Value", True, True, True, False, True).Value
52        IncludeAssetClasses = RangeFromSheet(shCreditUsage, "IncludeAssetClasses", False, True, False, False, False).Value
53        CurrenciesToInclude = RangeFromSheet(shConfig, "CurrenciesToInclude").Value
54        HedgeHorizon = GetHedgeHorizon()
55        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value
56        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False).Value
57        TimeStart = Now()
58        ComputerName = sEnvironmentVariable("ComputerName")

59        AnchorDate = GetItem(ModelBareBones, "AnchorDate")

60        N = RangeFromSheet(shScenarioDefinition, "ScenarioDefinition").Rows.Count - 1
61        PromptArray = sArrayStack("Number of time steps", N, _
              "ModelType", ModelType, _
              "NumMCPaths", Format(NumMCPaths, "###,###"), _
              "NumObservations", Format(NumObservations, "###,###"), _
              "FilterBy2", FilterBy2, _
              "Filter2Value", Filter2Value, _
              "IncludeAssetClasses", IncludeAssetClasses, _
              "CurrenciesToInclude", CurrenciesToInclude, _
              "HedgeHorizon", HedgeHorizon, _
              "OptionsStrategy:", "", _
              "ForwardsRatio", ForwardsRatio, _
              "PutRatio", PutRatio, _
              "CallRatio", CallRatio, _
              "PutStrikeOffset", PutStrikeOffset, _
              "CallStrikeOffset", CallStrikeOffset, _
              "StrategySwitchpoints", CStr(StrategySwitchPoints), _
              "AllocationByYear", AllocationByYear)

62        If TradesScaleFactor <> 1 Or LinesScaleFactor <> 1 Then
63            PromptArray = sArrayStack(PromptArray, _
                  "", "", _
                  "Morphing:", "", _
                  "TradesScaleFactor", TradesScaleFactor, _
                  "LinesScaleFactor", LinesScaleFactor)
64        End If

65        PromptArray = sReshape(PromptArray, sNRows(PromptArray) / 2, 2)
66        PromptArray = CleanUpPromptArray(PromptArray, True)
67        Prompt = "Run Scenario with the following inputs:" & vbLf & _
              sConcatenateStrings(sJustifyArrayOfStrings(PromptArray, "Calibri", 11, "  " & vbTab), vbLf)
              
68        If Not SilentMode Then
69            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, "Run Scenario", , , , , 600) <> vbOK Then Exit Function
70        End If

71        ShowFileInSnakeTail , True
72        Application.StatusBar = "The scenario is running. Progress is shown in the SnakeTail application."

73        ClearFutureTrades

74        With RangeFromSheet(shScenarioDefinition, "ScenarioDefinition")
75            ScenDefnHeaders = sArrayTranspose(.Rows(1).Value)
76            ScenarioDefinition = .offset(1).Resize(.Rows.Count - 1)
77        End With

          ' ScenarioResultsHeaders = sTokeniseString("1Y Traded,2Y Traded,3Y Traded,4Y Traded,5Y Traded,1Y capacity,2Y capacity,3Y capacity,4Y capacity,5Y capacity,USDPVAllTrades")
78        ScenarioResultsHeaders = sArrayStack(sArrayConcatenate(sIntegers(HedgeHorizon), "Y Traded"), _
              sArrayConcatenate(sIntegers(HedgeHorizon), "Y capacity"), _
              "USDPVAllTrades")

79        ScenarioResults = sReshape(Empty, sNRows(ScenarioDefinition), sNRows(ScenarioResultsHeaders))

          Dim cnr_10YC
          Dim cnr_10YT
          Dim cnr_1YC
          Dim cnr_1YT
          Dim cnr_2YC
          Dim cnr_2YT
          Dim cnr_3YC
          Dim cnr_3YT
          Dim cnr_4YC
          Dim cnr_4YT
          Dim cnr_5YC
          Dim cnr_5YT
          Dim cnr_6YC
          Dim cnr_6YT
          Dim cnr_7YC
          Dim cnr_7YT
          Dim cnr_8YC
          Dim cnr_8YT
          Dim cnr_9YC
          Dim cnr_9YT
          
          Dim cnr_PVUSD
          
80        cnr_1YT = ThrowIfError(sMatch("1Y Traded", ScenarioResultsHeaders))
81        cnr_2YT = ThrowIfError(sMatch("2Y Traded", ScenarioResultsHeaders))
82        cnr_3YT = ThrowIfError(sMatch("3Y Traded", ScenarioResultsHeaders))
83        cnr_4YT = ThrowIfError(sMatch("4Y Traded", ScenarioResultsHeaders))
84        cnr_5YT = ThrowIfError(sMatch("5Y Traded", ScenarioResultsHeaders))
85        If HedgeHorizon >= 6 Then cnr_6YT = ThrowIfError(sMatch("6Y Traded", ScenarioResultsHeaders))
86        If HedgeHorizon >= 7 Then cnr_7YT = ThrowIfError(sMatch("7Y Traded", ScenarioResultsHeaders))
87        If HedgeHorizon >= 8 Then cnr_8YT = ThrowIfError(sMatch("8Y Traded", ScenarioResultsHeaders))
88        If HedgeHorizon >= 9 Then cnr_9YT = ThrowIfError(sMatch("9Y Traded", ScenarioResultsHeaders))
89        If HedgeHorizon >= 10 Then cnr_10YT = ThrowIfError(sMatch("10Y Traded", ScenarioResultsHeaders))

90        cnr_1YC = ThrowIfError(sMatch("1Y capacity", ScenarioResultsHeaders))
91        cnr_2YC = ThrowIfError(sMatch("2Y capacity", ScenarioResultsHeaders))
92        cnr_3YC = ThrowIfError(sMatch("3Y capacity", ScenarioResultsHeaders))
93        cnr_4YC = ThrowIfError(sMatch("4Y capacity", ScenarioResultsHeaders))
94        cnr_5YC = ThrowIfError(sMatch("5Y capacity", ScenarioResultsHeaders))
95        If HedgeHorizon >= 6 Then cnr_6YC = ThrowIfError(sMatch("6Y capacity", ScenarioResultsHeaders))
96        If HedgeHorizon >= 7 Then cnr_7YC = ThrowIfError(sMatch("7Y capacity", ScenarioResultsHeaders))
97        If HedgeHorizon >= 8 Then cnr_8YC = ThrowIfError(sMatch("8Y capacity", ScenarioResultsHeaders))
98        If HedgeHorizon >= 9 Then cnr_9YC = ThrowIfError(sMatch("9Y capacity", ScenarioResultsHeaders))
99        If HedgeHorizon >= 10 Then cnr_10YC = ThrowIfError(sMatch("10Y capacity", ScenarioResultsHeaders))

100       cnr_PVUSD = ThrowIfError(sMatch("USDPVAllTrades", ScenarioResultsHeaders))

101       RangeFromSheet(shCreditUsage, "PortfolioAgeing").Value = 0
102       RangeFromSheet(shCreditUsage, "IncludeExtraTrades").Value = False
103       RangeFromSheet(shCreditUsage, "ExtraTradeAmounts").Value = 0
104       RangeFromSheet(shCreditUsage, "FilterBy1").Value = "Counterparty Parent"

          Dim cnd_AvEDSTraded
          Dim cnd_FxShock
          Dim cnd_FxVolShock
          Dim cnd_Months
          Dim cnd_ReplenishmentAmount
105       cnd_Months = ThrowIfError(sMatch("Months", ScenDefnHeaders))
106       cnd_FxShock = ThrowIfError(sMatch("FxShock", ScenDefnHeaders))
107       cnd_FxVolShock = ThrowIfError(sMatch("FxVolShock", ScenDefnHeaders))
108       cnd_ReplenishmentAmount = ThrowIfError(sMatch("ReplenishmentAmount", ScenDefnHeaders))
109       cnd_AvEDSTraded = sMatch("AvEDSTraded", ScenDefnHeaders)

110       AllocationsArray = ParseAllocation(AllocationByYear)

111       Application.GoTo shCreditUsage.Cells(1, 1)
112       For i = 1 To sNRows(ScenarioDefinition)
113           RefreshScreen
114           FxShock = ScenarioDefinition(i, cnd_FxShock)
115           FxVolShock = ScenarioDefinition(i, cnd_FxVolShock)
116           ThisDealDate = Application.WorksheetFunction.EDate(AnchorDate, ScenarioDefinition(i, cnd_Months))
117           PortfolioAgeing = (ThisDealDate - AnchorDate) / 365

118           AvEDSTraded = ScenarioDefinition(i, cnd_AvEDSTraded)

119           If StatusBarPrePrefix = "" Then
120               StatusBarPrefix = "RunScenario Month " & CStr(i) & " of " & CStr(N) & " "
121           Else
122               StatusBarPrefix = StatusBarPrePrefix & "Month " & CStr(i) & " of " & CStr(N) & " "
123           End If

              Const ThrowErrors As Boolean = True

124           WhoHasLines NumMCPaths, NumObservations, PortfolioAgeing, StatusBarPrefix, FxShock, FxVolShock, AllocationByYear, , , ThrowErrors
125           UpdateLinesHistory i, FxShock, FxVolShock, PortfolioAgeing

              'Madness having 10 variables rather than one vector variable. TODO fix this!
              Dim Amount10Y As Double
              Dim Amount1Y As Double
              Dim Amount2Y As Double
              Dim Amount3Y As Double
              Dim Amount4Y As Double
              Dim Amount5Y As Double
              Dim Amount6Y As Double
              Dim Amount7Y As Double
              Dim Amount8Y As Double
              Dim Amount9Y As Double
              
              Dim CCY1
              Dim CCY2
              Dim Ccy2Notional
              Dim Counterparties
              Dim MaturityLabels

              'Code below implements "Catch up" when we have previously not been able to get enough hedges on...
126           Amount1Y = 0: Amount2Y = 0: Amount3Y = 0: Amount4Y = 0: Amount5Y = 0: Amount6Y = 0: Amount7Y = 0: Amount8Y = 0: Amount9Y = 0: Amount10Y = 0
127           For j = 1 To i
128               Amount1Y = Amount1Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(1, 1)
129               Amount2Y = Amount2Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(2, 1)
130               Amount3Y = Amount3Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(3, 1)
131               Amount4Y = Amount4Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(4, 1)
132               Amount5Y = Amount5Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(5, 1)
133               If HedgeHorizon >= 6 Then Amount6Y = Amount6Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(6, 1)
134               If HedgeHorizon >= 7 Then Amount7Y = Amount7Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(7, 1)
135               If HedgeHorizon >= 8 Then Amount8Y = Amount8Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(8, 1)
136               If HedgeHorizon >= 9 Then Amount9Y = Amount9Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(9, 1)
137               If HedgeHorizon >= 10 Then Amount10Y = Amount10Y + ScenarioDefinition(j, cnd_ReplenishmentAmount) * AllocationsArray(10, 1)

138               If j < i Then
139                   Amount1Y = Amount1Y - ScenarioResults(j, cnr_1YT)
140                   Amount2Y = Amount2Y - ScenarioResults(j, cnr_2YT)
141                   Amount3Y = Amount3Y - ScenarioResults(j, cnr_3YT)
142                   Amount4Y = Amount4Y - ScenarioResults(j, cnr_4YT)
143                   Amount5Y = Amount5Y - ScenarioResults(j, cnr_5YT)
144                   If HedgeHorizon >= 6 Then Amount6Y = Amount6Y - ScenarioResults(j, cnr_6YT)
145                   If HedgeHorizon >= 7 Then Amount7Y = Amount7Y - ScenarioResults(j, cnr_7YT)
146                   If HedgeHorizon >= 8 Then Amount8Y = Amount8Y - ScenarioResults(j, cnr_8YT)
147                   If HedgeHorizon >= 9 Then Amount9Y = Amount9Y - ScenarioResults(j, cnr_9YT)
148                   If HedgeHorizon >= 10 Then Amount10Y = Amount10Y - ScenarioResults(j, cnr_10YT)
149               End If
150           Next j

151           AllocateTradesToBanks Amount1Y, Amount2Y, Amount3Y, Amount4Y, Amount5Y, _
                  Amount6Y, Amount7Y, Amount8Y, Amount9Y, Amount10Y, HedgeHorizon, _
                  Counterparties, CCY1, CCY2, Ccy2Notional, MaturityLabels, TotalsDone, _
                  TotalAvailableAfterAllocation, NumTradesDone

152           ScenarioResults(i, cnr_1YT) = TotalsDone(1, 1)
153           ScenarioResults(i, cnr_2YT) = TotalsDone(1, 2)
154           ScenarioResults(i, cnr_3YT) = TotalsDone(1, 3)
155           ScenarioResults(i, cnr_4YT) = TotalsDone(1, 4)
156           ScenarioResults(i, cnr_5YT) = TotalsDone(1, 5)
157           If HedgeHorizon >= 6 Then ScenarioResults(i, cnr_6YT) = TotalsDone(1, 6)
158           If HedgeHorizon >= 7 Then ScenarioResults(i, cnr_7YT) = TotalsDone(1, 7)
159           If HedgeHorizon >= 8 Then ScenarioResults(i, cnr_8YT) = TotalsDone(1, 8)
160           If HedgeHorizon >= 9 Then ScenarioResults(i, cnr_9YT) = TotalsDone(1, 9)
161           If HedgeHorizon >= 10 Then ScenarioResults(i, cnr_10YT) = TotalsDone(1, 10)

162           ScenarioResults(i, cnr_1YC) = TotalAvailableAfterAllocation(1, 1)
163           ScenarioResults(i, cnr_2YC) = TotalAvailableAfterAllocation(1, 2)
164           ScenarioResults(i, cnr_3YC) = TotalAvailableAfterAllocation(1, 3)
165           ScenarioResults(i, cnr_4YC) = TotalAvailableAfterAllocation(1, 4)
166           ScenarioResults(i, cnr_5YC) = TotalAvailableAfterAllocation(1, 5)
167           If HedgeHorizon >= 6 Then ScenarioResults(i, cnr_6YC) = TotalAvailableAfterAllocation(1, 6)
168           If HedgeHorizon >= 7 Then ScenarioResults(i, cnr_7YC) = TotalAvailableAfterAllocation(1, 7)
169           If HedgeHorizon >= 8 Then ScenarioResults(i, cnr_8YC) = TotalAvailableAfterAllocation(1, 8)
170           If HedgeHorizon >= 9 Then ScenarioResults(i, cnr_9YC) = TotalAvailableAfterAllocation(1, 9)
171           If HedgeHorizon >= 10 Then ScenarioResults(i, cnr_10YC) = TotalAvailableAfterAllocation(1, 10)

172           If NumTradesDone > 0 Then
173               TradeHeadroomInBillions = sSumOfNums(TotalAvailableAfterAllocation) / 1000000000#
174               SetOptionsStrategy TradeHeadroomInBillions, i, ForwardsRatio, PutRatio, CallRatio, _
                      PutStrikeOffset, CallStrikeOffset, StrategySwitchPoints, dForwardsRatio, dPutRatio, _
                      dCallRatio, dPutStrikeOffset, dCallStrikeOffset
175               ExecuteTrades ThisDealDate, Counterparties, CCY1, CCY2, Ccy2Notional, MaturityLabels, _
                      FxShock, AvEDSTraded, dForwardsRatio, dPutRatio, dCallRatio, dPutStrikeOffset, _
                      dCallStrikeOffset, ModelName, PortfolioAgeing, ModelBareBones
176           End If
              Dim PVUSD

177           PVUSD = sColumnSum(sColumnFromTable(RangeFromSheet(shWhoHasLines, "TheTable"), "PVUSD").Value)(1, 1)
178           ScenarioResults(i, cnr_PVUSD) = PVUSD
179       Next i

180       TimeEnd = Now()

181       RefreshScenarioResultsSheet shScenarioResults, ScenDefnHeaders, ScenarioDefinition, ScenarioResultsHeaders, _
              ScenarioResults, "Not yet saved to file", ShocksDerivedFrom, HistoryStart, HistoryEnd, BaseSpot, BaseVol, _
              ForwardsRatio, PutRatio, CallRatio, PutStrikeOffset, CallStrikeOffset, StrategySwitchPoints, _
              AllocationByYear, ModelType, NumMCPaths, NumObservations, FilterBy2, Filter2Value, _
              IncludeAssetClasses, CurrenciesToInclude, TradesScaleFactor, LinesScaleFactor, TimeStart, _
              TimeEnd, ComputerName, UseSpeedGrid, SpeedGridWidth, HighFxSpeed, LowFxSpeed, VaryGridWidth, _
              SpeedGridBaseVol, AnnualReplenishment, ScenarioDescription, HedgeHorizon

182       If SaveResults Then
183           FileName = RangeFromSheet(shConfig, "ScenarioResultsDirectory")
184           If Right(FileName, 1) <> "\" Then FileName = FileName & "\"
185           ThrowIfError sCreateFolder(FileName)
186           FileName = FileName + ScenarioDescriptionToFileName(RangeFromSheet(shScenarioDefinition, "ScenarioDescription"), "srf")
187           SaveScenarioResultsFile FileName, ""
188       End If

Cleanup:
189       InCleanUp = True
190       RangeFromSheet(shCreditUsage, "FxShock").Value = 1
191       RangeFromSheet(shCreditUsage, "FxVolShock").Value = 1
192       RangeFromSheet(shCreditUsage, "PortfolioAgeing").Value = 0
193       RangeFromSheet(shCreditUsage, "IncludeFutureTrades").Value = False
194       ClearoutResults
195       Application.StatusBar = False
196       RangeFromSheet(shScenarioResults, "TimeEnd").Value = Now()
197       SafeAppActivate SheetToActivate
198       RunScenario = IIf(CopyOfErr = "", "OK", CopyOfErr)
199       Exit Function
ErrHandler:
200       CopyOfErr = "#RunScenario (line " & CStr(Erl) & "): " & Err.Description & "!"
201       If SilentMode Then        'Calling from RunManyScenarios
202           RunScenario = CopyOfErr
203       Else
204           SomethingWentWrong "#RunScenario (line " & CStr(Erl) & "): " & Err.Description & "!"
205           If Not InCleanUp Then GoTo Cleanup
206       End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshScenarioResultsSheet
' Author    : Philip Swannell
' Date      : 24-Oct-2016
' Purpose   : Takes data generated by RunScenario and updates the ScenarioResults sheet
'             'The left 2 columns are "updated". The rest of the sheet, including the chart
'              are recreated from scratch.
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshScenarioResultsSheet(TargetSheet As Worksheet, ByVal ScenarioDefinitionHeaders, ScenarioDefinition, _
          ByVal ScenarioResultsHeaders, ScenarioResults, FileName As String, ShocksDerivedFrom As String, _
          HistoryStart As Long, HistoryEnd As Long, BaseSpot As Double, BaseVol As Double, ForwardsRatio As Variant, _
          PutRatio As Variant, CallRatio As Variant, PutStrikeOffset As Variant, CallStrikeOffset As Variant, _
          StrategySwitchPoints As Variant, AllocationByYear As String, ModelType As String, NumMCPaths As Long, _
          NumObservations As Long, FilterBy2 As String, Filter2Value As Variant, IncludeAssetClasses As String, _
          CurrenciesToInclude As String, TradesScaleFactor As Double, LinesScaleFactor As Double, _
          TimeStart As Variant, TimeEnd As Variant, ComputerName As String, UseSpeedGrid As Boolean, _
          SpeedGridWidth As Double, HighFxSpeed As Double, LowFxSpeed As Double, VaryGridWidth As Boolean, _
          SpeedGridBaseVol As Double, AnnualReplenishment As Double, ScenarioDescription As String, _
          HedgeHorizon As Long)

          Dim SPH As clsSheetProtectionHandler

          'Chart Legends
          Const L_HedgeCapacity = "Hedge Capacity ($ bln, right axis)"
          Const L_Spot = "EURUSD (left axis)"
          Const L_Vol = "3Y EURUSD Vol  (x 10, left axis)"
          Const L_HedgeCompetion = "Hedge completion ratio (%, right axis)"
          Const L_LineEx = "Line exhaustion ($ bln, right axis)"
          Const L_Time = "Time (months)"
          Const StartAddresss = "AA3"        'We paste in the data for plotting starting at this address
          Const TopLeftAddress = "D3"        'top left of chart goes here
          Const KPIAddress = "Q3"        'ScenarioPerformance goes here
          Dim IsHighRes As Boolean
          Dim NR As Long
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set ws = TargetSheet

3         Force2DArrayRMulti ScenarioDefinitionHeaders, ScenarioDefinition, ScenarioResultsHeaders, ScenarioResults
4         If sNRows(ScenarioDefinitionHeaders) = 1 Then
5             ScenarioDefinitionHeaders = sArrayTranspose(ScenarioDefinitionHeaders)
6         End If
7         If sNRows(ScenarioResultsHeaders) = 1 Then
8             ScenarioResultsHeaders = sArrayTranspose(ScenarioResultsHeaders)
9         End If
10        NR = sNRows(ScenarioDefinition)

          'Very basic error checking...
11        If sNRows(ScenarioResults) <> NR Then
12            Throw "ScenarioDefinition and ScenarioResults must have the same number of rows " & _
                  "but ScenarioDefinition has " & CStr(sNRows(ScenarioDefinition)) & _
                  " rows and ScenarioResults has " & CStr(sNRows(ScenarioResults)) & " rows."
13        End If

          'Read headers to get column numbers...
          Dim cnd_FxShock
          Dim cnd_FxVolShock
          Dim cnd_Months
          Dim cnd_ReplenishmentAmount
          Dim cnr_1YC
          Dim cnr_1YT
          Dim cnr_HHYC
          Dim cnr_HHYT
          
14        cnr_1YT = ThrowIfError(sMatch("1Y Traded", ScenarioResultsHeaders))
15        cnr_HHYT = ThrowIfError(sMatch(CStr(HedgeHorizon) & "Y Traded", ScenarioResultsHeaders))
16        cnr_1YC = ThrowIfError(sMatch("1Y capacity", ScenarioResultsHeaders))
17        cnr_HHYC = ThrowIfError(sMatch(CStr(HedgeHorizon) & "Y capacity", ScenarioResultsHeaders))

18        cnd_Months = ThrowIfError(sMatch("Months", ScenarioDefinitionHeaders))
19        cnd_FxShock = ThrowIfError(sMatch("FxShock", ScenarioDefinitionHeaders))
20        cnd_FxVolShock = ThrowIfError(sMatch("FxVolShock", ScenarioDefinitionHeaders))
21        cnd_ReplenishmentAmount = ThrowIfError(sMatch("ReplenishmentAmount", ScenarioDefinitionHeaders))

          'Calculate what we want to plot in the chart
          Dim ChooseVector
          Dim DatesForPlotting
          Dim HedgeCapacity
          Dim HedgeCompletionRatio
          Dim HistoryDates
          Dim HistorySpot
          Dim HistoryVol
          Dim LineExhaustionLevel
          Dim MaxMonth
          Dim SpotForPlotting
          Dim TargetTraded1toHH
          Dim TimeForPlotting
          Dim TotalTraded1toHH
          Dim VolForPlotting

22        HedgeCapacity = sRowSum(sSubArray(ScenarioResults, 1, cnr_1YT, , cnr_HHYC - cnr_1YT + 1))
23        HedgeCapacity = sArrayDivide(HedgeCapacity, 1000000000#)
24        TotalTraded1toHH = sRowSum(sSubArray(ScenarioResults, 1, cnr_1YT, , cnr_HHYT - cnr_1YT + 1))
25        TargetTraded1toHH = sSubArray(ScenarioDefinition, 1, cnd_ReplenishmentAmount, , 1)
26        HedgeCompletionRatio = sArrayMultiply(sArrayDivide(sPartialSum(TotalTraded1toHH), sPartialSum(TargetTraded1toHH)), 100)
27        LineExhaustionLevel = sReshape(10, NR, 1)
28        IsHighRes = LCase(ShocksDerivedFrom) = "history"

29        MaxMonth = ScenarioDefinition(sNRows(ScenarioDefinition), cnd_Months)
30        If IsHighRes Then
31            With RangeFromSheet(shHistoricalData, "TheDates")
32                HistoryDates = .Value
33                HistorySpot = .offset(, 1).Value
34                HistoryVol = .offset(, 2).Value
35            End With
36            ChooseVector = sArrayAnd(sArrayGreaterThanOrEqual(HistoryDates, HistoryStart), _
                  sArrayLessThanOrEqual(HistoryDates, HistoryEnd))
37            DatesForPlotting = sMChoose(HistoryDates, ChooseVector)
38            SpotForPlotting = sMChoose(HistorySpot, ChooseVector)
39            SpotForPlotting = sArrayMultiply(SpotForPlotting, BaseSpot / SpotForPlotting(1, 1))
40            VolForPlotting = sMChoose(HistoryVol, ChooseVector)
41            TimeForPlotting = sArraySubtract(DatesForPlotting, HistoryStart)

42            TimeForPlotting = sArrayMultiply(TimeForPlotting, MaxMonth / (HistoryEnd - HistoryStart))
43        Else
44            TimeForPlotting = sSubArray(ScenarioDefinition, 1, cnd_Months, , 1)
45            SpotForPlotting = sArrayMultiply(sSubArray(ScenarioDefinition, 1, cnd_FxShock, , 1), BaseSpot)
46            VolForPlotting = sArrayMultiply(sSubArray(ScenarioDefinition, 1, cnd_FxVolShock, , 1), BaseVol)
47        End If

48        VolForPlotting = sArrayMultiply(VolForPlotting, 10)        'To scale to fit on the chart

49        Set SPH = CreateSheetProtectionHandler(ws)

          'Clear out the sheet, apart from the left two columns
          Dim co As ChartObject
          Dim Res
50        For Each co In ws.ChartObjects
51            co.Delete
52        Next
          Dim N As Name
          Dim R As Range
53        For Each N In ws.Names
54            Set R = Nothing
55            On Error Resume Next
56            Set R = N.RefersToRange
57            On Error GoTo ErrHandler
58            If R Is Nothing Then
59                N.Delete
60            ElseIf R.Column > 2 Then
61                If R.Parent Is ws Then
62                    N.Delete
63                End If
64            End If
65        Next N
66        ws.Cells(1, 3).Resize(1, 1000).EntireColumn.Delete
67        Res = ws.UsedRange.Rows.Count        'resets used range

          'Paste in the Parameters to the left hand columns of the target sheet
68        With RangeFromSheet(ws, "FileName")
69            .Value = "'" & FileName
70            .offset(0, 2).Value = "'"
71        End With

72        RangeFromSheet(ws, "ShocksDerivedFrom").Value = ShocksDerivedFrom
73        RangeFromSheet(ws, "HistoryStart").Value = HistoryStart
74        RangeFromSheet(ws, "HistoryEnd").Value = HistoryEnd
75        RangeFromSheet(ws, "BaseSpot").Value = BaseSpot
76        RangeFromSheet(ws, "BaseVol").Value = BaseVol
77        RangeFromSheet(ws, "ForwardsRatio") = ForwardsRatio
78        RangeFromSheet(ws, "PutRatio").Value = PutRatio
79        RangeFromSheet(ws, "CallRatio").Value = CallRatio
80        RangeFromSheet(ws, "PutStrikeOffset").Value = PutStrikeOffset
81        RangeFromSheet(ws, "CallStrikeOffset").Value = CallStrikeOffset
82        RangeFromSheet(ws, "StrategySwitchPoints").Value = StrategySwitchPoints
83        RangeFromSheet(ws, "AllocationByYear").Value = AllocationByYear
84        RangeFromSheet(ws, "ModelType").Value = ModelType
85        RangeFromSheet(ws, "NumMCPaths").Value = NumMCPaths
86        RangeFromSheet(ws, "NumObservations").Value = NumObservations
87        RangeFromSheet(ws, "FilterBy2").Value = FilterBy2
88        RangeFromSheet(ws, "Filter2Value").Value = Filter2Value
89        RangeFromSheet(ws, "IncludeAssetClasses").Value = IncludeAssetClasses
90        RangeFromSheet(ws, "CurrenciesToInclude").Value = CurrenciesToInclude
91        RangeFromSheet(ws, "TradesScaleFactor").Value = TradesScaleFactor
92        RangeFromSheet(ws, "LinesScaleFactor").Value = LinesScaleFactor
93        RangeFromSheet(ws, "TimeStart").Value = TimeStart
94        RangeFromSheet(ws, "TimeEnd") = TimeEnd
95        RangeFromSheet(ws, "ComputerName").Value = ComputerName
96        RangeFromSheet(ws, "UseSpeedGrid").Value = UseSpeedGrid
97        RangeFromSheet(ws, "SpeedGridWidth").Value = IIf(UseSpeedGrid, SpeedGridWidth, Empty)
98        RangeFromSheet(ws, "HighFxSpeed").Value = IIf(UseSpeedGrid, HighFxSpeed, Empty)
99        RangeFromSheet(ws, "LowFxSpeed").Value = IIf(UseSpeedGrid, LowFxSpeed, Empty)
100       RangeFromSheet(ws, "VaryGridWidth").Value = IIf(UseSpeedGrid, VaryGridWidth, Empty)
101       RangeFromSheet(ws, "SpeedGridBaseVol").Value = IIf(UseSpeedGrid, SpeedGridBaseVol, Empty)
102       RangeFromSheet(ws, "AnnualReplenishment").Value = IIf(UseSpeedGrid, AnnualReplenishment, Empty)
103       RangeFromSheet(ws, "ScenarioDescription") = ScenarioDescription
          'Paste in the data for plotting...
          Dim StartCell As Range
          Dim TargetRange As Range
104       Set StartCell = ws.Range(StartAddresss)

          'Scenario Definition...

105       Set TargetRange = StartCell.Resize(sNRows(ScenarioDefinition) + 1, sNCols(ScenarioDefinition))
106       With TargetRange
107           .Value = sArrayStack(sArrayTranspose(ScenarioDefinitionHeaders), ScenarioDefinition)
108           .Rows(1).Font.Bold = True
109           AddGreyBorders .offset(0), True
110           ws.Names.Add "ScenarioDefinitionHeaders", .Rows(1)
111           ws.Names.Add "ScenarioDefinition", .offset(1).Resize(.Rows.Count - 1)
112           .Columns(cnd_ReplenishmentAmount).NumberFormat = "#,##0;[Red]-#,##0"
113           .Columns.AutoFit
114           .Cells(0, 1).Value = "ScenarioDefinition"
115           Set TargetRange = .Cells(1, .Columns.Count + 2)
116       End With

          'ScenarioResults...
117       Set TargetRange = TargetRange.Resize(sNRows(ScenarioResults) + 1, sNCols(ScenarioResults))
118       With TargetRange
119           .Value = sArrayStack(sArrayTranspose(ScenarioResultsHeaders), ScenarioResults)
120           .Rows(1).Font.Bold = True
121           AddGreyBorders .offset(0), True
122           ws.Names.Add "ScenarioResultsHeaders", .Rows(1)
123           ws.Names.Add "ScenarioResults", .offset(1).Resize(.Rows.Count - 1)
124           .NumberFormat = "#,##0;[Red]-#,##0"
125           .Columns.AutoFit
126           .Cells(0, 1).Value = "ScenarioResults"
127           .Cells(0, cnr_1YC) = "Capacity after allocation of trades"
128           Set TargetRange = .Cells(1, .Columns.Count + 2)
129       End With

          'HedgeCapacity etc...
130       Set TargetRange = TargetRange.Resize(sNRows(HedgeCapacity) + 1, 1)
131       With TargetRange
132           .Value = sArrayStack(sArrayTranspose(L_HedgeCapacity), HedgeCapacity)
133           AddGreyBorders .offset(0), True
134           ws.Names.Add "HedgeCapacity", .offset(1).Resize(.Rows.Count - 1)
135           .Columns.AutoFit
136           Set TargetRange = .Cells(1, .Columns.Count + 1)
137       End With
138       Set TargetRange = TargetRange.Resize(sNRows(HedgeCompletionRatio) + 1, 1)
139       With TargetRange
140           .Value = sArrayStack(sArrayTranspose(L_HedgeCompetion), HedgeCompletionRatio)
141           AddGreyBorders .offset(0), True
142           ws.Names.Add "HedgeCompletionRatio", .offset(1).Resize(.Rows.Count - 1)
143           .Columns.AutoFit
144           Set TargetRange = .Cells(1, .Columns.Count + 1)
145       End With
146       Set TargetRange = TargetRange.Resize(sNRows(LineExhaustionLevel) + 1, 1)
147       With TargetRange
148           .Value = sArrayStack(sArrayTranspose(L_LineEx), LineExhaustionLevel)
149           AddGreyBorders .offset(0), True
150           ws.Names.Add "LineExhaustionLevel", .offset(1).Resize(.Rows.Count - 1)
151           .Columns.AutoFit
152           Set TargetRange = .Cells(2, .Columns.Count + 2)
153       End With
154       Set TargetRange = TargetRange.Resize(sNRows(TimeForPlotting), 1)
155       With TargetRange
156           .Value = TimeForPlotting
157           AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
158           .Cells(0, 1) = L_Time
159           ws.Names.Add "TimeForPlotting", .offset(0)
160           .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
161           Set TargetRange = .Cells(1, .Columns.Count + 1)
162       End With
163       Set TargetRange = TargetRange.Resize(sNRows(SpotForPlotting), 1)
164       With TargetRange
165           .Value = SpotForPlotting
166           AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
167           .Cells(0, 1) = L_Spot
168           ws.Names.Add "SpotForPlotting", .offset(0)
169           .NumberFormat = "0.00"
170           .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
171           Set TargetRange = .Cells(1, .Columns.Count + 1)
172       End With
173       Set TargetRange = TargetRange.Resize(sNRows(VolForPlotting), 1)
174       With TargetRange
175           .Value = VolForPlotting
176           AddGreyBorders .offset(-1).Resize(.Rows.Count + 1), True
177           .Cells(0, 1) = L_Vol
178           ws.Names.Add "VolForPlotting", .offset(0)
179           .NumberFormat = "0.00"
180           .offset(-1).Resize(.Rows.Count + 1).Columns.AutoFit
181           Set TargetRange = .Cells(1, .Columns.Count + 2)
182       End With

          'Scenario Performance
          Dim ScenarioPerformance
183       ScenarioPerformance = sParseArrayString("{""Scenario Performance"",;""Number of stressed periods"",0;""Final Hedge Capacity"",0;""Final completion ratio"",0;""Worst completion ratio"",0}")
184       ScenarioPerformance(2, 2) = sArrayCount(sArrayLessThanOrEqual(HedgeCapacity, LineExhaustionLevel(1, 1)))
185       ScenarioPerformance(3, 2) = FirstElement(sTake(HedgeCapacity, -1))
186       ScenarioPerformance(4, 2) = FirstElement(sTake(HedgeCompletionRatio, -1))
187       ScenarioPerformance(5, 2) = FirstElement(sColumnMin(HedgeCompletionRatio))

188       With ws.Range(KPIAddress).Resize(sNRows(ScenarioPerformance), sNCols(ScenarioPerformance))
189           .Value = ScenarioPerformance
190           .Cells(1, 1).Font.Bold = True
191           .offset(1).Columns.AutoFit
192           .Columns(2).NumberFormat = "0.0"
193           .Cells(2, 2).NumberFormat = "General"
194           AddGreyBorders .offset(0), True
195           shScenarioResults.Names.Add "KPIs", .offset(0)
196       End With

          Dim MaximumScale As Double
          Dim MinimumScale As Double
197       MinimumScale = sMinOfNums(sArrayStack(SpotForPlotting, VolForPlotting))
198       MinimumScale = Application.WorksheetFunction.Floor(MinimumScale, 0.1)
199       MaximumScale = sMaxOfNums(sArrayStack(SpotForPlotting, VolForPlotting))
200       MaximumScale = Application.WorksheetFunction.Ceiling(MaximumScale, 0.1)

          Dim cht As Chart

          'Set up a new chart, using code that's compatible with Excel 2010 i.e. no FullSeriesCollection and no AddChart2
          ' Set cht = ws.Shapes.AddChart2(-1, xlXYScatterLinesNoMarkers, ws.Range(TopLeftAddress).Left, ws.Range(TopLeftAddress).Top, 590, 374.4).Chart
201       Set cht = ws.ChartObjects.Add(Left:=ws.Range(TopLeftAddress).Left, Top:=ws.Range(TopLeftAddress).Top, Width:=590, Height:=374.4).Chart
202       cht.ChartType = xlXYScatterLinesNoMarkers
203       cht.Parent.Placement = xlMove

204       With cht.SeriesCollection.NewSeries
205           .ChartType = xlXYScatterLinesNoMarkers
206           .Name = L_Spot
207           .xValues = "='" & ws.Name & "'!" & RangeFromSheet(ws, "TimeForPlotting").Address
208           .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "SpotForPlotting").Address
209       End With

210       With cht.SeriesCollection.NewSeries
211           .ChartType = xlXYScatterLinesNoMarkers
212           .Name = L_Vol
213           .xValues = "='" & ws.Name & "'!" & RangeFromSheet(ws, "TimeForPlotting").Address
214           .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "VolForPlotting").Address
215       End With

          Dim xValuesRange As Range
216       With RangeFromSheet(ws, "ScenarioDefinition")
217           Set xValuesRange = .Cells(1, cnd_Months).Resize(.Rows.Count)
218       End With

219       With cht.SeriesCollection.NewSeries
220           .ChartType = xlXYScatterLinesNoMarkers
221           .Name = L_HedgeCapacity
222           .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
223           .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "HedgeCapacity").Address
224       End With

225       With cht.SeriesCollection.NewSeries
226           .ChartType = xlXYScatterLinesNoMarkers
227           .Name = L_LineEx
228           .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
229           .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "LineExhaustionLevel").Address
230       End With

231       With cht.SeriesCollection.NewSeries
232           .ChartType = xlXYScatterLinesNoMarkers
233           .Name = L_HedgeCompetion
234           .xValues = "='" & ws.Name & "'!" & xValuesRange.Address
235           .Values = "='" & ws.Name & "'!" & RangeFromSheet(ws, "HedgeCompletionRatio").Address
236       End With

237       cht.SeriesCollection(1).AxisGroup = 1
238       cht.SeriesCollection(2).AxisGroup = 1
239       cht.SeriesCollection(3).AxisGroup = 2
240       cht.SeriesCollection(4).AxisGroup = 2
241       cht.SeriesCollection(5).AxisGroup = 2

242       cht.Axes(xlCategory).MinimumScale = 0
243       cht.Axes(xlCategory).MaximumScale = MaxMonth
244       cht.Axes(xlCategory).MajorUnit = 6
245       cht.Axes(xlValue, xlPrimary).MinimumScale = MinimumScale
246       cht.Axes(xlValue, xlPrimary).MaximumScale = MaximumScale

          'Have to change the ChartType late in the code...
247       cht.SeriesCollection(3).ChartType = xlColumnClustered
248       cht.SetElement (msoElementLegendBottom)
249       cht.SetElement (msoElementChartTitleAboveChart)
250       With cht.ChartTitle.Format.TextFrame2.TextRange.Font
251           .Bold = msoFalse
252           .Size = 14
253           .Italic = msoFalse
254       End With
255       cht.ChartTitle.Caption = "=ScenarioResults!ScenarioDescription"

256       With cht.Legend
257           .Left = 10
258           .Width = 570
259           .Height = 43.748
260           .Top = 324.651
261       End With
262       cht.PlotArea.Height = 290

          'This is strange. I have seen the chart end up with many (15!) series, but cannot replicate that mis-behaviour
          Dim i As Long
263       For i = cht.SeriesCollection.Count To 6 Step -1
264           cht.SeriesCollection(i).Delete
265       Next i

266       Exit Sub
ErrHandler:
267       Throw "#RefreshScenarioResultsSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestUpdateLinesHistory
' Author    : Philip Swannell
' Date      : 29-Sep-2016
' Purpose   : Test harness
' -----------------------------------------------------------------------------------------------------------------------
Sub TestUpdateLinesHistory()
1         On Error GoTo ErrHandler
2         Application.ScreenUpdating = False
3         UpdateLinesHistory 1, 1, 2, 3
4         UpdateLinesHistory 2, 4, 5, 6
5         UpdateLinesHistory 3, 7, 8, 9
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#TestUpdateLinesHistory (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UpdateLinesHistory
' Author    : Philip Swannell
' Date      : 17-Jun-2015
' Purpose   : Writes data from the WhoHasLines sheet into the LinesHistory sheet so that we
'             see more detail of what goes on during a Scenario.
' -----------------------------------------------------------------------------------------------------------------------
Sub UpdateLinesHistory(CallNumber As Long, FxShock, FxVolShock, PortfolioAgeing)
          Dim N As Name
          Dim Res As Variant
          Dim SourceData As Variant
          Dim SourceRange As Range
          Dim SPH As Object
          Dim TargetRange As Range

          Const StartRow = 4

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shLinesHistory)

3         If CallNumber = 1 Then
4             For Each N In shLinesHistory.Names
5                 N.Delete
6             Next
7             shLinesHistory.UsedRange.EntireColumn.Delete
8             Res = shLinesHistory.UsedRange.Rows.Count

9             With shLinesHistory.Cells(1, 1)
10                .Value = "Lines History"
11                .Font.Size = 22
12            End With

13            Set SourceRange = RangeFromSheet(shScenarioDefinition, "Parameters")
14            Set TargetRange = shLinesHistory.Cells(StartRow, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
15            With TargetRange
16                .Cells(0, 1).Value = "Parameters"
17                .Value = SourceRange.Value
18                .HorizontalAlignment = xlHAlignLeft
19                AddGreyBorders .offset(0), True
20                shLinesHistory.Names.Add "Parameters", .offset(0)
21                Set TargetRange = .Cells(.Rows.Count + 3, 1)
22            End With

23            SourceData = sArrayStack(sArrayRange("FxShock", FxShock), _
                  sArrayRange("FxVolShock", FxVolShock), _
                  sArrayRange("PortfolioAgeing", PortfolioAgeing))

24            Set TargetRange = TargetRange.Resize(sNRows(SourceData), sNCols(SourceData))
25            With TargetRange
26                .Cells(0, 1).Value = "Shocks" & CStr(CallNumber)
27                .Value = SourceData
28                .HorizontalAlignment = xlHAlignLeft
29                AddGreyBorders .offset(0), True
30                shLinesHistory.Names.Add "Shocks" & CStr(CallNumber), .offset(0)
31                Set TargetRange = .Cells(.Rows.Count + 3, 1)
32            End With

33            Set SourceRange = RangeFromSheet(shWhoHasLines, "TheTable")
34            Set TargetRange = TargetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
35            With TargetRange
36                .Cells(0, 1).Value = "TheTable" & CStr(CallNumber)
37                .Value = SourceRange.Value
38                .Columns(2).Resize(, .Columns.Count - 2).NumberFormat = "#,##0;[Red]-#,##0"
39                .HorizontalAlignment = xlHAlignLeft
40                AddGreyBorders .offset(0), True
41                shLinesHistory.Names.Add "TheTable" & CStr(CallNumber), .offset(0)
42                Application.Union(.offset(0), .Parent.Range("Parameters")).Columns.AutoFit
43            End With

44        Else

45            SourceData = sArrayStack(sArrayRange("FxShock", FxShock), _
                  sArrayRange("FxVolShock", FxVolShock), _
                  sArrayRange("PortfolioAgeing", PortfolioAgeing))

46            With shLinesHistory.Range("TheTable" & CStr(CallNumber - 1))
47                Set TargetRange = .Cells(-4, .Columns.Count + 2)
48            End With

49            Set TargetRange = TargetRange.Resize(sNRows(SourceData), sNCols(SourceData))
50            With TargetRange
51                .Cells(0, 1).Value = "Shocks" & CStr(CallNumber)
52                .Value = SourceData
53                .HorizontalAlignment = xlHAlignLeft
54                AddGreyBorders .offset(0), True
55                shLinesHistory.Names.Add "Shocks" & CStr(CallNumber), .offset(0)
56                Set TargetRange = .Cells(.Rows.Count + 3, 1)
57            End With

58            Set SourceRange = RangeFromSheet(shWhoHasLines, "TheTable")
59            Set TargetRange = TargetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
60            With TargetRange
61                .Cells(0, 1).Value = "TheTable" & CStr(CallNumber)
62                .Value = SourceRange.Value
63                .Columns(2).Resize(, .Columns.Count - 2).NumberFormat = "#,##0;[Red]-#,##0"
64                .HorizontalAlignment = xlHAlignLeft
65                AddGreyBorders .offset(0), True
66                shLinesHistory.Names.Add "TheTable" & CStr(CallNumber), .offset(0)
67                .Columns.AutoFit
68            End With
69        End If

70        Exit Sub
ErrHandler:
71        Throw "#UpdateLinesHistory (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CleanUpPromptArray
' Author    : Philip Swannell
' Date      : 14-Jun-2015
' Purpose   : Remove elements from the PromptArray that are "standard" so that the user can focus on what's unusual.
' -----------------------------------------------------------------------------------------------------------------------
Function CleanUpPromptArray(ByVal PromptArray, Optional WithAreYouSure As Boolean = False)
          Dim ChooseVector
          Dim i As Long
          Dim MatchRes
          Dim NullElements
          Dim OptionsStrategyLabels
          Dim OptionsStrategyValues
          Dim Result

          'Remove dull values so user more likely to see what matters
1         NullElements = sArrayStack(sArrayRange("FxShock", "1"), _
              sArrayRange("FxVolShock", "1"), _
              sArrayRange("PortfolioAgeing", "0"), _
              sArrayRange("FilterBy2", "None"), _
              sArrayRange("Filter2Value", "None"), _
              sArrayRange("TradesScaleFactor", "1"), _
              sArrayRange("LinesScaleFactor", "1"), _
              sArrayRange("Calculate Trade headroom", "False"), _
              sArrayRange("Calculate Fx headroom", "False"), _
              sArrayRange("IncludeExtraTrades", "False"), _
              sArrayRange("IncludeFutureTrades", "False"), _
              sArrayRange("PortfolioAgeing", "0"), _
              sArrayRange("IncludeAssetClasses", "Rates and Fx"), _
              sArrayRange("NumMCPaths", "255"))

2         On Error GoTo ErrHandler

3         If sVLookup("FilterBy2", PromptArray) = "None" Then
4             MatchRes = sMatch("Filter2Value", sSubArray(PromptArray, 1, 1, , 1))
5             If IsNumber(MatchRes) Then PromptArray(MatchRes, 2) = "None"
6         End If

7         For i = 1 To sNRows(PromptArray)
8             If Not IsEmpty(PromptArray(i, 2)) Then
9                 PromptArray(i, 2) = CStr(PromptArray(i, 2))
10            End If
11        Next i

12        OptionsStrategyLabels = sArrayStack("ForwardsRatio", "PutRatio", "CallRatio", _
              "PutStrikeOffset", "CallStrikeOffset", "StrategySwitchPoints", "OptionsStrategy:")
13        OptionsStrategyValues = sVLookup(OptionsStrategyLabels, PromptArray)
14        If sArraysIdentical(sTake(OptionsStrategyValues, 6), sArrayStack(1, 0, 0, 0, 0, "")) Then
15            NullElements = sArrayStack(NullElements, sArrayRange(OptionsStrategyLabels, OptionsStrategyValues))
16        End If

17        ChooseVector = sArrayNot(sArrayIsNumber(sMultiMatch(PromptArray, NullElements, False)))
18        Result = sMChoose(PromptArray, ChooseVector)

19        If WithAreYouSure Then
20            For i = 1 To sNRows(Result)
21                Select Case Result(i, 1)
                      Case "FxShock", "FxVolShock", "PortfolioAgeing", "TradesScaleFactor", "LinesScaleFactor"
22                        Result(i, 2) = CStr(Result(i, 2)) & "                               <-- ARE YOU SURE?"
23                End Select

24            Next i
25        End If

26        CleanUpPromptArray = Result

27        Exit Function
ErrHandler:
28        Throw "#CleanUpPromptArray (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsAscendingArrayOfNaturals
' Author    : Philip Swannell
' Date      : 20-Oct-2016
' Purpose   : Code to test the "Months" column entered in a scenario - must be ascending array of non-negative integers
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsAscendingArrayOfNaturals(Months)
          Const MonthsError = "Months must be a 1-column array of positive whole numbers in ascending order"
1         Force2DArrayRMulti Months
2         If sNCols(Months) <> 1 Then Throw MonthsError

          Dim i As Long

3         On Error GoTo ErrHandler
4         For i = 1 To sNRows(Months)
5             If Not IsNumber(Months(i, 1)) Then
6                 Throw MonthsError
7             ElseIf CLng(Months(i, 1)) <> Months(i, 1) Then
8                 Throw MonthsError
9             ElseIf Months(i, 1) < 0 Then
10                Throw MonthsError
11            ElseIf i > 1 Then
12                If Months(i, 1) < Months(i - 1, 1) Then
13                    Throw MonthsError
14                End If
15            End If
16        Next i
17        Exit Function
ErrHandler:
18        Throw "#IsAscendingArrayOfNaturals (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShockFromHistory
' Author    : Philip Swannell
' Date      : 20-Oct-2016
' Purpose   : Derive the shocks to Fx spot and vol from a historical period, data for which
'             is saved on the HistoricalData sheet. Note how we shock FxVol so that the
'             absolute levels for the scenario match absolute levels in the
'             historic period. By contrast, for Fx we shock so that month on month changes
'             match that seen in the historic period.
' -----------------------------------------------------------------------------------------------------------------------
Function ShockFromHistory(HistoryStart As Long, Months As Variant, IsSpot As Boolean, Optional FxVolToday As Double)

          Dim i As Long
          Dim InterpolatedValues
          Dim MaxDate
          Dim MaxDateNeeded As Long
          Dim MaxMonths As Long
          Dim MinDate
          Dim xArrayAscending As Variant
          Dim xValues As Variant
          Dim yArray As Variant

1         On Error GoTo ErrHandler
2         IsAscendingArrayOfNaturals Months        'throws if there's bad data

3         xArrayAscending = shHistoricalData.Range("TheDates").Value
4         If IsSpot Then
5             yArray = shHistoricalData.Range("TheDates").offset(, 1).Value
6         Else
7             yArray = shHistoricalData.Range("TheDates").offset(, 2).Value
8         End If

9         MinDate = xArrayAscending(1, 1)
10        MaxDate = sTake(xArrayAscending, -1)(1, 1)

11        If Not IsSpot Then
12            If FxVolToday = 0 Then
13                Throw "If IsSpot is FALSE then FxVolToday must be provided so that generated " & _
                      "shocks shock today's vol to the vol levels in the Historical period"
14            End If
15        End If

16        If HistoryStart < MinDate Then
17            Throw "Historical Data is only available starting from " & Format(MinDate, "dd-mmm-yyyy") + _
                  " but a " & IIf(IsSpot, "spot rate", "volatility") & _
                  " is required for " & Format(HistoryStart, "dd-mmm-yyyy")
18        End If

19        MaxMonths = FirstElementOf(sColumnMax(Months))
20        MaxDateNeeded = Application.WorksheetFunction.EDate(HistoryStart, MaxMonths)
21        If MaxDateNeeded > MaxDate Then
22            Throw "Historical Data is only available to " & _
                  Format(MaxDate, "dd-mmm-yyyy") & " but a " & IIf(IsSpot, "spot rate", "volatility") + _
                  " is required for " & Format(MaxDateNeeded, "dd-mmm-yyyy")
23        End If

24        xValues = sReshape(0, sNRows(Months), 1)
25        For i = 1 To sNRows(Months)
26            xValues(i, 1) = Application.WorksheetFunction.EDate(HistoryStart, Months(i, 1))
27        Next i

28        If IsSpot Then
29            InterpolatedValues = sInterp(xArrayAscending, yArray, sArrayStack(HistoryStart, xValues), "FlatFromLeft")
30            ShockFromHistory = sArrayDivide(sDrop(InterpolatedValues, 1), InterpolatedValues(1, 1))
31        Else
32            InterpolatedValues = sInterp(xArrayAscending, yArray, xValues, "FlatFromLeft")
33            ShockFromHistory = sArrayDivide(InterpolatedValues, FxVolToday)
34        End If

35        On Error GoTo ErrHandler

36        Exit Function
ErrHandler:
37        Throw "#ShockFromHistory (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AvEDSTradedFromHistory
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : Models Airbus's "Speed Grid" hedging strategy to determine what their monthly
'            total hedge amounts would have been in a historic scenario. Return is either
'           two columns (Verbose = FALSE) of ReplenishmentAmounts and average spot fx (AVEDSTraded)
'           or else a multi-column array with headers for use to populate the SpeedGridDrillDown sheet
' -----------------------------------------------------------------------------------------------------------------------
Function AvEDSTradedFromHistory(HistoryStart As Long, Months As Variant, SpeedGridWidth As Double, HighFxSpeed, _
          LowFxSpeed, VaryGridWidth As Boolean, SpeedGridBaseVol As Double, RebaseSpotToStartAt As Double, _
          AnnualReplenishment As Double, Optional Verbose As Boolean = False)
          
          Dim HistoricDates
          Dim HistoricSpots
          Dim HistoricVols
          Dim i As Long
          Dim InterpRes
          Dim j As Long
          Dim k As Long
          Dim MaxHistory As Long
          Dim MaxMonth As Long
          Dim MinHistory As Long
          Dim MinMonth As Long
          Dim NumMonths As Long
          Dim RebaseFactor As Double
          Dim FromIndexes, ToIndexes        'Provides the coordinates for subarray-ing into the historic data for each monthly period
          'Dim the columns of a sensible verbose return - i.e. a return that walks through the calculation
          Dim AvEDSTraded
          Dim AverageSpeed
          Dim HedgingStates
          Dim HiBarrier
          Dim HiMinusLo
          Dim IndexOfLastDay As Long
          Dim LoBarrier
          Dim MonthlyDates
          Dim NDBetween
          Dim NDGH
          Dim NDLL
          Dim NumBizDays
          Dim PartialSums
          Dim ReplenishmentAmounts
          Dim SoMFx
          Dim SoMFxRebased
          Dim SoMVol

1         On Error GoTo ErrHandler
2         NumMonths = sNRows(Months)

3         If Not (sArraysIdentical(Months, sIntegers(NumMonths))) Then
4             Throw "Months must be integers from 1 to " & CStr(NumMonths)
5         End If

6         HistoricDates = shHistoricalData.Range("TheDates").Value2
7         HistoricSpots = shHistoricalData.Range("TheDates").offset(, 1).Value
8         HistoricVols = shHistoricalData.Range("TheDates").offset(, 2).Value
9         MonthlyDates = sReshape(0, NumMonths + 1, 1)

10        MonthlyDates(1, 1) = HistoryStart
11        For i = 1 To NumMonths
12            MonthlyDates(i + 1, 1) = Application.WorksheetFunction.EDate(HistoryStart, Months(i, 1))
13        Next

14        MaxMonth = Months(NumMonths, 1)
15        MinMonth = HistoryStart
16        MaxHistory = HistoricDates(sNRows(HistoricDates), 1)
17        MinHistory = HistoricDates(1, 1)

18        If MinMonth < MinHistory Then
19            Throw "Historical Data is only available starting from " & Format(MinHistory, "dd-mmm-yyyy") + _
                  " but data is required for " & Format(MinMonth, "dd-mmm-yyyy")
20        End If

21        If MaxMonth > MaxHistory Then
22            Throw "Historical Data is only available to " & _
                  Format(MaxHistory, "dd-mmm-yyyy") & " but data is required for " & Format(MaxMonth, "dd-mmm-yyyy")
23        End If

24        NumBizDays = sReshape(0, NumMonths, 1)

25        InterpRes = ThrowIfError(sInterp(sArrayStack(0, MonthlyDates, 1000000), sGrid(0, NumMonths + 2, NumMonths + 3), HistoricDates, "FlatFromLeft", "NN"))

          Dim Zeros
26        Zeros = sReshape(0, NumMonths, 1)

27        NumBizDays = Zeros
28        FromIndexes = Zeros
29        ToIndexes = Zeros
30        SoMFx = Zeros
31        SoMVol = Zeros
32        NDGH = Zeros
33        NDLL = Zeros
34        NDBetween = Zeros

35        For i = 1 To sNRows(HistoricDates)
36            If IsNumber(InterpRes(i, 1)) Then
37                k = InterpRes(i, 1)
38                If k = 0 Then
                      'nothing to do
39                ElseIf k <= NumMonths Then
40                    IndexOfLastDay = i
41                    If NumBizDays(k, 1) = 0 Then
42                        FromIndexes(k, 1) = i
43                        If k > 1 Then
44                            ToIndexes(k - 1, 1) = FromIndexes(k, 1) - 1
45                        End If
46                    End If
47                    NumBizDays(k, 1) = NumBizDays(k, 1) + 1
48                Else
49                    Exit For
50                End If
51            End If
52        Next
53        ToIndexes(NumMonths, 1) = IndexOfLastDay

54        For i = 1 To NumMonths
55            If HistoricDates(FromIndexes(i, 1), 1) = MonthlyDates(i, 1) Then
56                SoMFx(i, 1) = HistoricSpots(FromIndexes(i, 1), 1)
57                SoMVol(i, 1) = HistoricVols(FromIndexes(i, 1), 1)
58            Else
59                SoMFx(i, 1) = HistoricSpots(FromIndexes(i, 1) - 1, 1)
60                SoMVol(i, 1) = HistoricVols(FromIndexes(i, 1) - 1, 1)
61            End If
62        Next i

63        RebaseFactor = RebaseSpotToStartAt / SoMFx(1, 1)
64        SoMFxRebased = sArrayMultiply(SoMFx, RebaseFactor)
65        If VaryGridWidth Then
              Dim SpeedGridWidths
66            SpeedGridWidths = sArrayMultiply(SpeedGridWidth, sArrayDivide(SoMVol, SpeedGridBaseVol))
67            HiBarrier = sArrayAdd(SoMFxRebased, sArrayDivide(SpeedGridWidths, 2))
68            LoBarrier = sArraySubtract(SoMFxRebased, sArrayDivide(SpeedGridWidths, 2))
69        Else
70            HiBarrier = sArrayAdd(SoMFxRebased, SpeedGridWidth / 2)
71            LoBarrier = sArraySubtract(SoMFxRebased, SpeedGridWidth / 2)
72        End If
73        HiMinusLo = sArraySubtract(HiBarrier, LoBarrier)

74        For i = 1 To NumMonths
75            For j = FromIndexes(i, 1) To ToIndexes(i, 1)
76                If HistoricSpots(j, 1) * RebaseFactor > HiBarrier(i, 1) Then
77                    NDGH(i, 1) = NDGH(i, 1) + 1
78                ElseIf HistoricSpots(j, 1) * RebaseFactor < LoBarrier(i, 1) Then
79                    NDLL(i, 1) = NDLL(i, 1) + 1
80                Else
81                    NDBetween(i, 1) = NDBetween(i, 1) + 1
82                End If
83            Next j
84        Next i

85        AverageSpeed = sArrayAdd(sArrayMultiply(NDGH, HighFxSpeed), sArrayMultiply(NDLL, LowFxSpeed), NDBetween)
86        AverageSpeed = sArrayDivide(AverageSpeed, NumBizDays)

          Const State_SpeedGridOperating = "Speed Grid Operating"
          Const State_SpeedGridStopDuringMonth = "Hedging Stops during Month"
          Const State_SpeedGridNoHedging = "No Hedging"
          Const State_SpeedGridCatchUp = "Catch up needed"
87        HedgingStates = Zeros
88        PartialSums = Zeros

89        For i = 1 To NumMonths
90            PartialSums(i, 1) = AverageSpeed(i, 1)
91            If i Mod 12 <> 1 Then
92                PartialSums(i, 1) = PartialSums(i, 1) + PartialSums(i - 1, 1)
93            End If
94        Next i

95        For i = 1 To NumMonths
96            If i Mod 12 = 1 Then
97                If PartialSums(i, 1) < 12 Then
98                    HedgingStates(i, 1) = State_SpeedGridOperating
99                Else
100                   HedgingStates(i, 1) = State_SpeedGridStopDuringMonth
101               End If
102           ElseIf i Mod 12 = 0 Then
103               If PartialSums(i - 1, 1) > 12 Then
104                   HedgingStates(i, 1) = State_SpeedGridNoHedging
105               Else
106                   HedgingStates(i, 1) = State_SpeedGridCatchUp
107               End If
108           Else
109               If PartialSums(i, 1) < 12 Then
110                   HedgingStates(i, 1) = State_SpeedGridOperating
111               ElseIf PartialSums(i, 1) >= 12 And PartialSums(i - 1, 1) < 12 Then
112                   HedgingStates(i, 1) = State_SpeedGridStopDuringMonth
113               Else
114                   HedgingStates(i, 1) = State_SpeedGridNoHedging
115               End If
116           End If
117       Next i

118       ReplenishmentAmounts = Zeros
119       AvEDSTraded = Zeros

          Dim PriceDetails
          Dim TradeDetails
120       If Verbose Then
121           PriceDetails = Zeros: TradeDetails = Zeros
122       End If

          Dim DailyAmount
          Dim DailyAmounts
          Dim DailyPrices
          Dim MonthlyAmount As Double
123       MonthlyAmount = AnnualReplenishment / 12

124       For i = 1 To NumMonths
125           DailyAmount = MonthlyAmount / NumBizDays(i, 1)
126           DailyPrices = sReshape(0, CLng(NumBizDays(i, 1)), 1)
127           DailyAmounts = DailyPrices

128           Select Case HedgingStates(i, 1)
                  Case State_SpeedGridNoHedging
129                   ReplenishmentAmounts(i, 1) = 0
130                   AvEDSTraded(i, 1) = 0
131               Case State_SpeedGridOperating, State_SpeedGridStopDuringMonth
132                   k = 0
133                   For j = FromIndexes(i, 1) To ToIndexes(i, 1)
134                       k = k + 1
135                       If HistoricSpots(j, 1) * RebaseFactor > HiBarrier(i, 1) Then
136                           DailyPrices(k, 1) = HistoricSpots(j, 1) * RebaseFactor
137                           DailyAmounts(k, 1) = DailyAmount * HighFxSpeed
138                       ElseIf HistoricSpots(j, 1) * RebaseFactor < LoBarrier(i, 1) Then
139                           DailyPrices(k, 1) = HistoricSpots(j, 1) * RebaseFactor
140                           DailyAmounts(k, 1) = DailyAmount * LowFxSpeed
141                       Else
142                           DailyPrices(k, 1) = HistoricSpots(j, 1) * RebaseFactor
143                           DailyAmounts(k, 1) = DailyAmount
144                       End If
145                   Next j
146                   If HedgingStates(i, 1) = State_SpeedGridStopDuringMonth Then
                          Dim TargetForMonth
147                       If i Mod 12 = 1 Then
148                           TargetForMonth = AnnualReplenishment
149                       Else
150                           TargetForMonth = AnnualReplenishment * (12 - PartialSums(i - 1, 1)) / 12
151                           If TargetForMonth < 0 Then
152                               Throw "Assertion failed - negative value for 'TargetForMonth'"
153                           End If
154                       End If
155                       DailyAmounts = TruncatePartialSum(DailyAmounts, TargetForMonth)
156                   End If
157                   ReplenishmentAmounts(i, 1) = FirstElementOf(sColumnSum(DailyAmounts))
158                   AvEDSTraded(i, 1) = FirstElementOf(sArrayDivide(sColumnSum(sArrayMultiply(DailyAmounts, DailyPrices)), ReplenishmentAmounts(i, 1)))
159               Case State_SpeedGridCatchUp
                      'No speed grid use
160                   TargetForMonth = AnnualReplenishment * (12 - PartialSums(i - 1, 1)) / 12
161                   k = 0
162                   For j = FromIndexes(i, 1) To ToIndexes(i, 1)
163                       k = k + 1
164                       DailyPrices(k, 1) = HistoricSpots(j, 1) * RebaseFactor
165                       DailyAmounts(k, 1) = TargetForMonth / NumBizDays(i, 1)
166                   Next j
167                   ReplenishmentAmounts(i, 1) = TargetForMonth
168                   AvEDSTraded(i, 1) = FirstElementOf(sColumnSum(DailyPrices)) / NumBizDays(i, 1)
169           End Select

170           If Verbose Then
171               For j = 1 To sNRows(DailyPrices)
172                   DailyPrices(j, 1) = Format(DailyPrices(j, 1), "0.00000")
173                   DailyAmounts(j, 1) = Format(DailyAmounts(j, 1) / 1000000, "0.000")
174               Next j
175               PriceDetails(i, 1) = sConcatenateStrings(DailyPrices, ";")
176               TradeDetails(i, 1) = sConcatenateStrings(DailyAmounts, ";")
177           End If
178       Next i

179       If Verbose Then
              Dim Headers

180           Headers = sArrayTranspose(sTokeniseString("Period" & vbLf & "Start,Period" & vbLf & "End,Num" & vbLf & _
                  "business" & vbLf & "days,Start of" & vbLf & "Month Spot,Start of Month" & vbLf & "Spot Rebased," & _
                  "Start of" & vbLf & "Month Vol,Hi Barrier,Low Barrier,Hi - Low,Num" & vbLf & "Days > Hi,Num" & vbLf & _
                  "Days < Low," & "Num Days between" & vbLf & "Hi and Low,Average hedging speed if" & vbLf & _
                  "Speed Grid operating," & "Annual Partial Sum" & vbLf & "of AHSISGO,Speed Grid State,Replenishment" & _
                  vbLf & "Amount," & "Average spot weighted" & vbLf & "by trade size,Daily trades: spot levels," & _
                  "Daily trades: Sizes (USD millions)"))

181           AvEDSTradedFromHistory = sArrayStack(Headers, sArrayRange(sDrop(MonthlyDates, -1), _
                  sDrop(MonthlyDates, 1), NumBizDays, SoMFx, SoMFxRebased, SoMVol, HiBarrier, _
                  LoBarrier, HiMinusLo, NDGH, NDLL, NDBetween, AverageSpeed, _
                  PartialSums, HedgingStates, ReplenishmentAmounts, AvEDSTraded, _
                  PriceDetails, TradeDetails))

182       Else
183           AvEDSTradedFromHistory = sArrayRange(ReplenishmentAmounts, AvEDSTraded)
184       End If

185       Exit Function
ErrHandler:
186       AvEDSTradedFromHistory = "#AvEDSTradedFromHistory (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TruncatePartialSum
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : takes a column of numbers and, reading down the column, when the partial
'             sum reached a limit sets the numbers to zero, or (in at most one case) to
'             the amount to take the sum of the output to limit.
'Example ColOfNums = 10,20,30,20
'            Limit = 37
'           Return = 10,20,7,0
' -----------------------------------------------------------------------------------------------------------------------
Function TruncatePartialSum(ColOfNums, Limit)
          Dim PS
1         On Error GoTo ErrHandler
2         PS = sArrayStack(0, sPartialSum(ColOfNums))
3         PS = sArrayMin(PS, Limit)
4         TruncatePartialSum = sArraySubtract(sDrop(PS, 1), sDrop(PS, -1))
5         Exit Function
ErrHandler:
6         Throw "#TruncatePartialSum (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub QuickTest()
1         On Error GoTo ErrHandler
2         RefreshScenarioDefinition True
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#QuickTest (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ParseAllocation
' Author    : Philip Swannell
' Date      : 21-Oct-2016
' Purpose   : Parses the AllocationByYear string to a 5 by 1 array of numbers summing to
'            one. Also validate the input and throws an error if input is invalid.
' -----------------------------------------------------------------------------------------------------------------------
Function ParseAllocation(ByRef AllocationByYear As String, Optional FromDialog As Boolean, Optional Normalise As Boolean = True)
          Dim ErrorString As String
          Dim FoundNonZero As Boolean

          Dim Example As String
          Dim HH As Long
          Dim i As Long
          Dim NR As Long
          Dim Res

1         On Error GoTo ErrHandler

          'For brevity...
2         Select Case LCase(AllocationByYear)
              Case "1y"
3                 AllocationByYear = "1"
4             Case "2y", "3y", "4y", "5y", "6y", "7y", "8y", "9y", "10y"
5                 AllocationByYear = sConcatenateStrings(sArrayStack(sReshape(0, CLng(Left(AllocationByYear, Len(AllocationByYear) - 1) - 1), 1), 1), ":")
6         End Select

7         HH = GetHedgeHorizon()

8         Example = "0:0:1:1:1"

9         ErrorString = "AllocationByYear must be a delimited string with up to " & CStr(HH) & " tokens denoting " & _
              "the proportions traded at 1 to " & CStr(HH) & " years. Delimiter can be comma, space or colon. " & _
              "Example '" & Example & "' for equal hedging in 3, 4 and 5 years." & vbLf & vbLf & _
              "The string was rejected because "

10        AllocationByYear = Replace(AllocationByYear, ",", ":")
11        AllocationByYear = Replace(AllocationByYear, " ", ":")
12        AllocationByYear = Replace(AllocationByYear, ";", ":")
13        Res = sTokeniseString(AllocationByYear, ":")
          
          'For backward-compatibility
14        NR = sNRows(Res)
15        If NR < HH Then
16            Res = sArrayStack(Res, sReshape(0, HH - NR, 1))
17        End If

18        If sNRows(Res) <> HH Then Throw ErrorString & "there were " & CStr(sNRows(Res)) & " token(s) instead of " & CStr(HH)
19        For i = 1 To HH
20            If Not IsNumeric(Res(i, 1)) Then Throw ErrorString & "element " & CStr(i) & " was '" & CStr(Res(i, 1)) & "' rather than a number"
21            Res(i, 1) = CDbl(Res(i, 1))
22            If Res(i, 1) < 0 Then Throw ErrorString & "element " & CStr(i) & " was negative, which is not allowed"
23            If Res(i, 1) <> 0 Then FoundNonZero = True
24        Next i
25        If Not FoundNonZero Then Throw ErrorString & "all of the allocations were zero"

26        If Normalise Then
27            ParseAllocation = sArrayDivide(Res, sColumnSum(Res))
28        Else
29            ParseAllocation = Res
30        End If

31        Exit Function
ErrHandler:
32        If FromDialog Then
33            Throw Err.Description
34        Else
35            Throw "#ParseAllocation (line " & CStr(Erl) & "): " & Err.Description & "!"
36        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshScenarioDefinition
' Author    : Philip Swannell
' Date      : 20-Oct-2016
' Purpose   : Applies cell formatting and also, for Scenarios that have ShocksDerivedFromHistory,
'             sets the FxShocks and FxVolShocks. For SpeedGrid scenarios calls AvEDSTradedFromHistory
'             to populate the replenishment amounts and the AvEDSTraded columns. Aim is to make the
'             ScenarioDefinition sheet easy to use and make our old ScenarioDefinitionFileGenerator
'             redundant. :-)
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshScenarioDefinition(DoFullValidation As Boolean, Optional DoDrillDown As Boolean = False)
          Dim AllocationByYear
          Dim AnchorDate
          Dim AnnualReplenishment As Double
          Dim c As Range
          Dim FxShocks
          Dim FxVolShocks
          Dim HighFxSpeed As Double
          Dim HistoryEnd As Variant
          Dim HistoryStart As Variant
          Dim LowFxSpeed As Double
          Dim ModelType As String
          Dim Months As Variant
          Dim rngSD As Range
          Dim ShocksDerivedFrom As String
          Dim SpeedGridBaseVol As Double
          Dim SpeedGridWidth As Double
          Dim SPH As clsSheetProtectionHandler
          Dim SpotToday As Variant
          Dim SUH As clsScreenUpdateHandler
          Dim UseHistory As Boolean
          Dim UseSpeedGrid As Boolean
          Dim VaryGridWidth As Boolean
          Dim VolToday As Double
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         Set ws = shScenarioDefinition

3         Set SUH = CreateScreenUpdateHandler()
4         Set SPH = CreateSheetProtectionHandler(ws)
5         Set rngSD = RangeFromSheet(ws, "ScenarioDefinition")

6         ModelType = MT_HW

          'Check ShocksDerivedFrom...
7         ShocksDerivedFrom = RangeFromSheet(ws, "ShocksDerivedFrom", False, True, False, False, False).Value
8         Select Case ShocksDerivedFrom
              Case "History", "Custom"
9                 UseHistory = LCase(ShocksDerivedFrom) = "history"
10            Case Else
11                Throw "ShocksDervidedFrom must be either 'History' or 'Custom'"
12        End Select

          Dim CallRatio
          Dim CallStrikeOffset
          Dim ForwardsRatio
          Dim PutRatio
          Dim PutStrikeOffset
          Dim StrategySwitchPoints
13        ForwardsRatio = RangeFromSheet(ws, "ForwardsRatio", True, True, False, False, False)
14        PutRatio = RangeFromSheet(ws, "PutRatio", True, True, False, False, False)
15        CallRatio = RangeFromSheet(ws, "CallRatio", True, True, False, False, False)
16        PutStrikeOffset = RangeFromSheet(ws, "PutStrikeOffset", True, True, False, False, False)
17        CallStrikeOffset = RangeFromSheet(ws, "CallStrikeOffset", True, True, False, False, False)
18        StrategySwitchPoints = RangeFromSheet(ws, "StrategySwitchPoints", True, True, False, True, False)

          'Check Months
19        Months = sColumnFromTable(rngSD, "Months").Value2
20        IsAscendingArrayOfNaturals Months
21        AllocationByYear = RangeFromSheet(ws, "AllocationByYear").Value

22        UseSpeedGrid = RangeFromSheet(ws, "UseSpeedGrid", False, False, True, False, False)

23        If DoDrillDown Then If Not UseSpeedGrid Then Throw "UseSpeedGrid must be TRUE for drill-down"

24        If UseSpeedGrid Then
25            SpeedGridWidth = RangeFromSheet(ws, "SpeedGridWidth", True, False, False, False, False)
26            If SpeedGridWidth < 0 Then Throw "SpeedGridWidth must be positive. e.g. 0.05 for 5 cents width"
27            HighFxSpeed = RangeFromSheet(ws, "HighFxSpeed", True, False, False, False, False)
28            If HighFxSpeed < 0 Then
29                Throw "HighFxSpeed must be positive or zero e.g. 0.5 for hedging at 50% of the standard rate when spot is above the monthly high barrier"
30            End If
31            LowFxSpeed = RangeFromSheet(ws, "LowFxSpeed", True, False, False, False, False)
32            If LowFxSpeed < 0 Then
33                Throw "LowFxSpeed must be positive or zero e.g. 1.5 for hedging at 150% of the standard rate when spot is below the monthly low barrier"
34            End If
35            VaryGridWidth = RangeFromSheet(ws, "VaryGridWidth", False, False, True, False, False)
36            SpeedGridBaseVol = RangeFromSheet(ws, "SpeedGridBaseVol", True, False, False, False, False)
37            If SpeedGridBaseVol < 0 Then Throw "SpeedGridBaseVol must be positve"
38            AnnualReplenishment = RangeFromSheet(ws, "AnnualReplenishment", True, False, False, False, False)
39            If AnnualReplenishment < 0 Then Throw "AnnualReplenishment must be positive"
40            If ShocksDerivedFrom <> "History" Then Throw "When UseSpeedGid is TRUE, ShocksDerivedFrom must be 'History'", True
41        End If

42        If DoFullValidation Then
              'When calling this method from the worksheet change event we don't want to validate _
               everything since the user is in the process of editing, but when calling at the start _
               of a scenario run we want to do much more validation...

43            ParseAllocation CStr(AllocationByYear)

44            SetOptionsStrategy 100, 1, ForwardsRatio, _
                  PutRatio, _
                  CallRatio, _
                  PutStrikeOffset, _
                  CallStrikeOffset, _
                  StrategySwitchPoints, _
                  0, 0, 0, 0, 0

45            For Each c In sColumnFromTable(rngSD, "FxShock").Cells
46                If Not IsPositive(c.Value) Then
47                    Throw "All FxShocks must be positive numbers but value at cell " & Replace(c.Address, "$", "") & " is not"
48                End If
49            Next c
50            For Each c In sColumnFromTable(rngSD, "FxVolShock").Cells
51                If Not IsPositive(c.Value) Then
52                    Throw "All FxVolShocks must be positive numbers but value at cell " & Replace(c.Address, "$", "") & " is not"
53                End If
54            Next c
55            For Each c In sColumnFromTable(rngSD, "ReplenishmentAmount").Cells
56                If Not IsPositiveOrZero(c.Value) Then
57                    Throw "All ReplenishmentAmounts must be positive numbers or zero but value at cell " & Replace(c.Address, "$", "") & " is not"
58                End If
59            Next c
60            For Each c In sColumnFromTable(rngSD, "AvEDSTraded").Cells
61                If Not IsPositiveOrZero(c.Value) Then
62                    Throw "All AvEDSTradeds must be zero or positive or zero but value at cell " & Replace(c.Address, "$", "") & " is not"
63                End If
64            Next c

65        End If

66        HistoryStart = RangeFromSheet(ws, "HistoryStart", True, False, False, False, False).Value
67        HistoryEnd = Application.WorksheetFunction.EDate(HistoryStart, Months(sNRows(Months), 1))

          'Get data that allows calculation of the shocks
68        OpenMarketWorkbook True, False
69        JuliaLaunchForCayley
70        BuildModelsInJulia False, 1, 1

71        AnchorDate = RangeFromMarketDataBook("Config", "AnchorDate").Value2
72        VolToday = GetItem(gModel_CM, "EURUSD3YVol")
73        SpotToday = GetItem(gModel_CM, "EURUSD")

          'Get these before we make any changes to the sheet, since invalid start or end date will cause error
74        If ShocksDerivedFrom = "History" Then
75            FxShocks = ShockFromHistory(CLng(HistoryStart), Months, True)
76            FxVolShocks = ShockFromHistory(CLng(HistoryStart), Months, False, VolToday)
77        End If

78        If UseSpeedGrid Then
              Dim AvEDSResult
              Dim SpeedGridFailed As Boolean
79            AvEDSResult = AvEDSTradedFromHistory(CLng(HistoryStart), Months, SpeedGridWidth, HighFxSpeed, _
                  LowFxSpeed, VaryGridWidth, SpeedGridBaseVol, CDbl(SpotToday), AnnualReplenishment, DoDrillDown)
80            If DoDrillDown Then
81                DrillDownIntoSpeedGrid AvEDSResult
82                Exit Sub
83            End If
84            SpeedGridFailed = sIsErrorString(AvEDSResult)
85        ElseIf UseHistory Then
86            AvEDSResult = AvEDSTradedFromHistory(CLng(HistoryStart), Months, 10, 1, 1, False, _
                  0.1, CDbl(SpotToday), 20000000000#, DoDrillDown)
87            SpeedGridFailed = sIsErrorString(AvEDSResult)
88        End If

          'First pass at cell formatting...
89        With rngSD
90            AddGreyBorders .offset(0), True
91            .HorizontalAlignment = xlHAlignCenter
92            CayleyFormatAsInput .offset(1).Resize(.Rows.Count - 1)
93            With sColumnFromTable(.offset(0), "Months")
94                .NumberFormat = "General"
95                .Locked = True
96                .Font.Color = g_Col_GreyText
97                .Value = sIntegers(.Rows.Count)
98            End With
99            With sColumnFromTable(.offset(0), "ReplenishmentAmount")
100               .NumberFormat = "#,##0;[Red]-#,##0"
101           End With
102           With sColumnFromTable(.offset(0), "AvEDSTraded")
103               .NumberFormat = "0.0000"
104           End With
105           With sColumnFromTable(.offset(0), "FxShock")
106               .NumberFormat = "0.0000"
107           End With
108           With sColumnFromTable(.offset(0), "FxVolShock")
109               .NumberFormat = "0.0000"
110           End With
111           .offset(.Rows.Count).Resize(120).Clear
112       End With

          'Second pass of cell formatting
113       If ShocksDerivedFrom = "History" Then
114           With sColumnFromTable(rngSD, "FxShock")
115               .Value = FxShocks
116               .Locked = True
117               .Font.Color = g_Col_GreyText
118           End With

119           With sColumnFromTable(rngSD, "FxVolShock")
120               .Value = FxVolShocks
121               .Locked = True
122               .Font.Color = g_Col_GreyText
123           End With
124           CayleyFormatAsInput RangeFromSheet(ws, "HistoryStart")
125       Else
126           With RangeFromSheet(ws, "HistoryStart")
127               .Locked = True
128               .Font.Color = g_Col_GreyText
129           End With
130       End If

131       With sColumnFromTable(rngSD, "ReplenishmentAmount")
132           If UseSpeedGrid And Not SpeedGridFailed Then
133               .Value = sSubArray(AvEDSResult, 1, 1, , 1)
134               .Locked = True
135               .Font.Color = g_Col_GreyText
136           Else
137               .Locked = False
138               CayleyFormatAsInput .offset(0)
139           End If
140       End With

141       With sColumnFromTable(rngSD, "AvEDSTraded")
142           .Locked = True
143           .Font.Color = g_Col_GreyText
144           If UseSpeedGrid And Not SpeedGridFailed Then
145               .Value = sSubArray(AvEDSResult, 1, 2, , 1)
146           ElseIf UseHistory And Not SpeedGridFailed Then
147               .Value = sSubArray(AvEDSResult, 1, 2, , 1)
148           Else
149               .Value = 0
150           End If
151       End With

152       RangeFromSheet(ws, "BaseSpot").Value = SpotToday
153       RangeFromSheet(ws, "BaseVol").Value = VolToday
154       RangeFromSheet(ws, "HistoryEnd").Value = HistoryEnd

155       rngSD.Columns.AutoFit

156       With Range(RangeFromSheet(ws, "SpeedGridWidth"), RangeFromSheet(ws, "AnnualReplenishment"))
157           If UseSpeedGrid Then
158               .Locked = False
159               CayleyFormatAsInput .offset(0)
160           Else
161               .Locked = True
162               .Font.Color = g_Col_GreyText
163           End If

164       End With
165       If Not VaryGridWidth Then
166           With RangeFromSheet(ws, "SpeedGridBaseVol")
167               .Locked = True
168               .Font.Color = g_Col_GreyText
169           End With
170       End If

          Dim LinesScaleFactor As Double
          Dim TradesScaleFactor As Double
171       TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False)
172       LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False)

          Dim ScenarioDescription
173       ScenarioDescription = DescribeScenario(sColumnFromTable(rngSD, "FxShock"), _
              sColumnFromTable(rngSD, "ReplenishmentAmount"), ShocksDerivedFrom, HistoryStart, _
              HistoryEnd, ForwardsRatio, PutRatio, CallRatio, PutStrikeOffset, CallStrikeOffset, _
              StrategySwitchPoints, AllocationByYear, UseSpeedGrid, SpeedGridWidth, HighFxSpeed, _
              LowFxSpeed, VaryGridWidth, AnnualReplenishment, TradesScaleFactor, _
              LinesScaleFactor)

174       RangeFromSheet(ws, "ScenarioDescription").Value = ScenarioDescription

          'Need to have one volatile formula on the sheet so that Shift F9 can be captured in the worksheet calculate event
175       RangeFromSheet(ws, "CellWithFormula").Formula = "=IF(RAND()<0.5,"""","""")"

          Dim Res
176       Res = ws.UsedRange.Rows.Count

177       If SpeedGridFailed Then Throw AvEDSResult

178       Exit Sub
ErrHandler:
179       Throw "#RefreshScenarioDefinition (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Function IsPositive(x As Variant) As Boolean
1         On Error GoTo ErrHandler
2         If IsNumber(x) Then IsPositive = x > 0
3         Exit Function
ErrHandler:
4         IsPositive = False
End Function

Private Function IsPositiveOrZero(x As Variant) As Boolean
1         On Error GoTo ErrHandler
2         If IsNumber(x) Then IsPositiveOrZero = x >= 0
3         Exit Function
ErrHandler:
4         IsPositiveOrZero = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DrillDownIntoSpeedGrid
' Author    : Philip Swannell
' Date      : 26-Oct-2016
' Purpose   : Called from the menu
' -----------------------------------------------------------------------------------------------------------------------
Sub DrillDownIntoSpeedGrid(ArrayToShow)
1         On Error GoTo ErrHandler
2         ThisWorkbook.Unprotect
          Dim N As Name
          Dim SourceRange As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim ws As Worksheet

3         Set ws = shSpeedGridDrillDown
4         Set SPH = CreateSheetProtectionHandler(ws)
5         Set SUH = CreateScreenUpdateHandler()
6         ws.Visible = xlSheetVisible

7         ws.UsedRange.EntireColumn.Delete
8         ws.UsedRange.EntireRow.Delete

9         For Each N In ws.Names
10            N.Delete
11        Next N

12        With ws.Cells(1, 1)
13            .Value = "Speed Grid drill-down"
14            .Font.Size = 22
15        End With

16        Set SourceRange = Range(RangeFromSheet(shScenarioDefinition, "FileName").Cells(0.1), _
              RangeFromSheet(shScenarioDefinition, "AnnualReplenishment"))

17        Set TargetRange = ws.Cells(3, 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)

18        SourceRange.Copy
19        TargetRange.PasteSpecial xlPasteAll
20        TargetRange.offset(2).Columns.AutoFit
21        TargetRange.Locked = True
22        TargetRange.Font.ColorIndex = xlColorIndexAutomatic

23        Set TargetRange = ws.Cells(2, 4).Resize(sNRows(ArrayToShow), sNCols(ArrayToShow))

24        With TargetRange
25            .Value = sArrayExcelString(ArrayToShow)
26            .Columns.ColumnWidth = 200
27            .Columns.AutoFit
28            .HorizontalAlignment = xlHAlignCenter
29            AddGreyBorders .offset(0), True
30            .Columns(1).NumberFormat = "dd-mmm-yyyy"        'PeriodStart
31            .Columns(2).NumberFormat = "dd-mmm-yyyy"        'PeriodEnd
32            .Columns(3).NumberFormat = "General"        'Num business days
33            .Columns(4).NumberFormat = "0.0000"        'Start of Month Spot
34            .Columns(5).NumberFormat = "0.0000"        'Start of Month Spot Rebased
35            .Columns(6).NumberFormat = "0.00%"        'Start of Month Vol
36            .Columns(7).NumberFormat = "0.0000"        'Hi Barrier
37            .Columns(8).NumberFormat = "0.0000"        'Low Barrier
38            .Columns(9).NumberFormat = "0.0000"        'Hi - Low
39            .Columns(10).NumberFormat = "General"        'NumDays > Hi
40            .Columns(11).NumberFormat = "General"        'Num Days < Low
41            .Columns(12).NumberFormat = "General"        'Num Days between Hi and Low
42            .Columns(13).NumberFormat = "0.0000"        'Average hedging speed if Speed Grid operating
43            .Columns(14).NumberFormat = "0.0000"        'Annual Partial Sum of AHSISGO
44            .Columns(15).NumberFormat = "General"        'Speed Grid State
45            .Columns(16).NumberFormat = "#,##0;[Red]-#,##0"        'Replenishment Amount
46            .Columns(17).NumberFormat = "0.0000"        'Average spot weighted by trade size
47            .Columns(18).NumberFormat = "General"        'Daily trades: spot levels
48            .Columns(19).NumberFormat = "General"        'Daily trades: Sizes (USD millions)
49            .Columns.AutoFit
50            With .Rows(1)
51                .Rows(1).WrapText = True
52                .VerticalAlignment = xlVAlignCenter
53            End With
54        End With

55        Application.GoTo ws.Cells(1, 1)

56        Exit Sub
ErrHandler:
57        Throw "#DrillDownIntoSpeedGrid (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DescribeScenario
' Author    : Philip Swannell
' Date      : 27-Oct-2016
' Purpose   : Returns a text string briefly describing a scenario
' -----------------------------------------------------------------------------------------------------------------------
Function DescribeScenario(FxShocks, ReplenishmentAmounts, ShocksDerivedFrom As String, _
          HistoryStart As Variant, HistoryEnd, ForwardsRatio, PutRatio, CallRatio, _
          PutStrikeOffset, CallStrikeOffset, StrategySwitchPoints, AllocationByYear, _
          UseSpeedGrid As Boolean, SpeedGridWidth, HighFxSpeed, LowFxSpeed, _
          VaryGridWidth As Boolean, AnnualReplenishment As Double, _
          TradesScaleFactor As Double, LinesScaleFactor As Double)

1         On Error GoTo ErrHandler
2         Force2DArrayR ReplenishmentAmounts

          Dim FirstPart As String
          Dim FirstReplenishmentAmount
          Dim LastReplenishmentAmount
          Dim NMonths As Long
          Dim SwitchOnTime As Boolean
3         NMonths = sNRows(FxShocks)

          'First part gives information about annual replenishment amount
4         If UseSpeedGrid Then
5             FirstPart = "$" & Format(AnnualReplenishment / 1000000000#, "0") & " bn pa "
6         Else
              Dim ReplenishmentsByYear
7             ReplenishmentsByYear = sColumnSumByChunks(ReplenishmentAmounts, 12)
8             FirstReplenishmentAmount = ReplenishmentsByYear(1, 1)
9             If NMonths < 12 Then
10                FirstReplenishmentAmount = FirstReplenishmentAmount / NMonths * 12
11            End If
12            LastReplenishmentAmount = ReplenishmentsByYear(sNRows(ReplenishmentsByYear), 1)
13            If NMonths Mod 12 <> 0 Then
14                LastReplenishmentAmount = LastReplenishmentAmount / (NMonths Mod 12) * 12
15            End If

16            If FirstReplenishmentAmount = LastReplenishmentAmount Then
17                FirstPart = "$" & Format(FirstReplenishmentAmount / 1000000000#, "0") & " bn pa "
18            Else
19                FirstPart = "$" & Format(FirstReplenishmentAmount / 1000000000#, "0") + _
                      " - $" & Format(LastReplenishmentAmount / 1000000000#, "0") & " bn pa "
20            End If
21        End If
22        FirstPart = FirstPart + AbbreviateAllocationByYear(CStr(AllocationByYear)) & " "

          'Second part about trading strategy
          Dim i As Long
          Dim NumStrategies
          Dim SecondPart As String
          Dim TokenisedSwitchPoints

23        If IsEmpty(StrategySwitchPoints) Then
24            NumStrategies = 1
25        Else
26            TokenisedSwitchPoints = sTokeniseString(CStr(StrategySwitchPoints))
27            Force2DArray TokenisedSwitchPoints

28            NumStrategies = sNRows(TokenisedSwitchPoints) + 1

29            If LCase(TokenisedSwitchPoints(1, 1)) = LCase("SwitchOnTime") Then
30                SwitchOnTime = True
31                NumStrategies = NumStrategies - 1
32            End If
33        End If
          Dim ThisCallRatio
          Dim ThisCallStrikeOffset
          Dim ThisForwardsRatio
          Dim ThisPutRatio
          Dim ThisPutStrikeOffset

34        For i = 1 To NumStrategies
35            If i > 1 Then
36                SecondPart = SecondPart & "/"
37            End If
38            If NumStrategies = 1 Then
39                ThisForwardsRatio = CStr(ForwardsRatio)
40                ThisPutRatio = CStr(PutRatio)
41                ThisCallRatio = CStr(CallRatio)
42                ThisPutStrikeOffset = CStr(PutStrikeOffset)
43                ThisCallStrikeOffset = CStr(CallStrikeOffset)
44            Else
45                ThisForwardsRatio = sTokeniseString(CStr(ForwardsRatio))(i, 1)
46                ThisPutRatio = sTokeniseString(CStr(PutRatio))(i, 1)
47                ThisCallRatio = sTokeniseString(CStr(CallRatio))(i, 1)
48                ThisPutStrikeOffset = sTokeniseString(CStr(PutStrikeOffset))(i, 1)
49                ThisCallStrikeOffset = sTokeniseString(CStr(CallStrikeOffset))(i, 1)
50            End If
51            If ThisForwardsRatio <> "0" Then
52                SecondPart = SecondPart + Format(CDbl(ThisForwardsRatio), "00%") & " fwds "
53            End If
54            If CDbl(ThisCallRatio) = -CDbl(ThisPutRatio) Then
55                If ThisCallRatio <> "0" Then
56                    SecondPart = SecondPart + Format(CDbl(ThisCallRatio), "00%") & " collars "
57                End If
58            Else
59                If ThisCallRatio <> "0" Then
60                    SecondPart = SecondPart + Format(CDbl(ThisCallRatio), "00%") & " calls "
61                End If
62                If ThisPutRatio <> "0" Then
63                    SecondPart = SecondPart + Format(CDbl(ThisPutRatio), "00%") & " puts "
64                End If
65            End If
66            If i > 1 Then
67                If SwitchOnTime Then
68                    SecondPart = SecondPart & "(time>=" & sTokeniseString(CStr(StrategySwitchPoints))(i, 1) & "months) "
69                Else
70                    SecondPart = SecondPart & "(HR<" & sTokeniseString(CStr(StrategySwitchPoints))(i - 1, 1) & ") "
71                End If
72            End If
73        Next i

          'Describe speed grid
          Dim ThirdPart As String

74        If UseSpeedGrid And VaryGridWidth Then
75            ThirdPart = "Dynamic Speed Grid(" & CStr(SpeedGridWidth) & "," & Format(HighFxSpeed, "00%") & "," & Format(LowFxSpeed, "00%") & ") "
76        ElseIf UseSpeedGrid Then
77            ThirdPart = "Speed Grid(" & CStr(SpeedGridWidth) & "," & Format(HighFxSpeed, "00%") & "," & Format(LowFxSpeed, "00%") & ") "
78        End If

          'Describe Shocks
          Dim FourthPart As String
79        If LCase(ShocksDerivedFrom) = "history" Then
80            FourthPart = "Path: " & Format(HistoryStart, "mmm yy") & " - " & Format(HistoryEnd, "mmm yy") & " "
81        Else
82            If NMonths Mod 12 = 0 Then
83                FourthPart = CStr(NMonths / 12) & "Y Custom Path "
84            Else
85                FourthPart = CStr(sNRows(FxShocks)) & "M Custom Path "
86            End If
87        End If

          Dim FifthPart As String
          Dim SixthPart As String
          'Describe morphing
88        If TradesScaleFactor <> 1 Then
89            FifthPart = "Trades x " & CStr(TradesScaleFactor) & " "
90        End If
91        If LinesScaleFactor <> 1 Then
92            SixthPart = "Lines x " & CStr(LinesScaleFactor) & " "
93        End If

94        DescribeScenario = Trim(FirstPart + SecondPart + ThirdPart + FifthPart + SixthPart + FourthPart)

95        Exit Function
ErrHandler:
96        DescribeScenario = "#DescribeScenario (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AbbreviateAllocationByYear
' Author     : Philip Swannell
' Date       : 15-Feb-2022
' Purpose    : Abbreviate the AllocationsByYear string for use within a file name. e.g. "1:1:1:1:0:0:0:0" becomes "1-4Y"
' -----------------------------------------------------------------------------------------------------------------------
Function AbbreviateAllocationByYear(ByVal AllocationByYear As String)
          Dim AllocsArray01
          Dim ChooseVector
          Dim CRret
          Dim Res As String

          'We only look at hedge or not hedge in a year, don't try to describe 1:2 ratios
1         On Error GoTo ErrHandler
2         AllocationByYear = sConcatenateStrings(sArrayIf(sArrayEquals(sTokeniseString(AllocationByYear, ":"), "0"), "0", "1"), ":")

3         AllocsArray01 = sArrayIf(sArrayEquals(sTokeniseString(AllocationByYear, ":"), "0"), "0", "1")
4         CRret = sCountRepeats(AllocsArray01, "CFT")
5         If sNRows(CRret) <= 2 Then
6             ChooseVector = sArrayEquals("1", sSubArray(CRret, 1, 1, , 1))
7             If sArrayCount(ChooseVector) = 1 Then
8                 CRret = sMChoose(CRret, ChooseVector)
9                 Res = CStr(CRret(1, 2)) & "-" & CStr(CRret(1, 3)) & "Y"
10            End If
11        End If

12        If Res = "" Then
13            If sAll(sArrayEquals(AllocsArray01, 0)) Then
14                Res = "No hedging!"
15            Else
16                Res = sConcatenateStrings(sMChoose(sIntegers(sNRows(AllocsArray01)), sArrayEquals("1", AllocsArray01))) & "Y"
17            End If
18        End If
19        AbbreviateAllocationByYear = Res
20        Exit Function
ErrHandler:
21        Throw "#AbbreviateAllocationByYear (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

