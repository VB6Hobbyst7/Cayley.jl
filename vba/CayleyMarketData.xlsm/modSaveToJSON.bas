Attribute VB_Name = "modSaveToJSON"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestSaveDataFromMarketWorkbookToFile
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Test harness...
' -----------------------------------------------------------------------------------------------------------------------
Sub TestSaveDataFromMarketWorkbookToFile()
          Const FileName = "c:\temp\foo.json"
1         On Error GoTo ErrHandler

2         tic
3         ThrowIfError SaveDataFromMarketWorkbookToFile(ThisWorkbook, FileName, sArrayStack("EUR", "USD", "GBP"), "EUR", sArrayStack(gSELF, "BARC_GB_LON"), 3)
4         toc "SaveDataFromMarketWorkbookToFile"

5         ShowFileInTexteditor FileName
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#TestSaveDataFromMarketWorkbookToFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveDataFromMarketWorkbookToFile
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : For an input set of currencies, credits and inflation indices, saves to a JSON file all of the market data needed by our R code
'              If there's an error the function returns an error string - so we may call via Application.Run
' Parameters :
'  MarketWb           : A "market data workbook" such as this one.
'  FileName           : Name, with path, of the file to be created. Existing file of that name is overwritten.
'  CCys               : Columm array of currencies - 3-letter ISO codes as usual.
'  Numeraire          : The Numeraire of the multi currency Hull White Model that R will use.
'  Credits            : A column array of credits - must be a subset of the left column of data held on the Credit sheet of the MDWb
'  WithHistoricalFxVol: For project Cayley only, should the file also contain historical Fx vols?
'  FxVolHandling      : 1 (or True for back compat) write both FxVol and FxVolHistoric, in keys such as FxVol_EURUSD and FxVolHistoric_EURUSD
'                       2 (or False) write only FxVol, in keys such as FxVol_EURUSD
'                       3  write only FxVolHistoric, but in keys such as FxVol_EURUSD

'  InflationIndices   : A column array of inflation indices. Must be a subset of the sheet names in this workbook which are inflation sheets -
' -----------------------------------------------------------------------------------------------------------------------
Function SaveDataFromMarketWorkbookToFile(MarketWb As Workbook, FileName As String, ByVal CCys As Variant, Numeraire As String, Optional ByVal Credits As Variant, _
          Optional FxVolHandling As Variant = 2, Optional ByVal InflationIndices As Variant)

          Dim Ccy As Variant
          Dim DataToWrite As String
          Dim DCT As New Dictionary
          Dim InflationIndex As Variant
          Dim OldSaved As Boolean
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         OldSaved = ThisWorkbook.Saved

3         Set SUH = CreateScreenUpdateHandler()

4         If Not (IsEmpty(InflationIndices) Or IsMissing(InflationIndices)) Then
5             Force2DArray InflationIndices
6             For Each InflationIndex In InflationIndices
7                 If Not IsInCollection(MarketWb.Worksheets, CStr(InflationIndex)) Then Throw "The market data workbook must have a sheet called '" + CStr(InflationIndex) + "' with market data for the inflation index " + CStr(InflationIndex), True
8                 If Not IsInflationSheet(ThisWorkbook.Worksheets(InflationIndex)) Then Throw "Unexpected error: Method 'IsInflationSheet' returns FALSE for the worksheet '" + InflationIndex + "' of workbook '" + MarketWb.Name + "'"
                  Dim ThisBaseCurrency As String
9                 ThisBaseCurrency = sVLookup("BaseCurrency", MarketWb.Worksheets(CStr(InflationIndex)).Range("Parameters").Value)
10                If Not IsNumber(sMatch(ThisBaseCurrency, CCys)) Then
11                    CCys = sArrayStack(CCys, ThisBaseCurrency)
12                End If
13            Next
14        End If

15        HideUnhideSheets MarketWb, sArrayStack(Numeraire, CCys), Numeraire, InflationIndices

16        DCT.Add "AnchorDate", sFormatDate(RangeFromSheet(shConfig, "AnchorDate", True, False, False, False, False).Value2, "YYYY-MM-DD")
17        DCT.Add "Currencies", To1D(CCys)
18        If IsEmpty(Credits) Or IsMissing(Credits) Or IsNull(Credits) Then
19            Credits = Null
20            DCT.Add "Credits", Credits
21        Else
22            Force2DArray Credits
23            DCT.Add "Credits", To1D(Credits)
24        End If

25        If IsEmpty(InflationIndices) Or IsMissing(InflationIndices) Or IsNull(InflationIndices) Then InflationIndices = Null Else InflationIndices = To1D(InflationIndices)
26        DCT.Add "Inflations", InflationIndices
27        DCT.Add "Numeraire", RangeFromSheet(shConfig, "Numeraire", False, True, False, False, False).Value
28        DCT.Add "CollateralCcy", RangeFromSheet(shConfig, "CollateralCcy", False, True, False, False, False).Value
29        DCT.Add "SigmaStep", RangeFromSheet(shConfig, "SigmaStep", True, False, False, False, False).Value
30        DCT.Add "TStar", RangeFromSheet(shConfig, "TStar", True, False, False, False, False).Value
31        DCT.Add "HWRevert", RangeFromSheet(shConfig, "HWRevert", True, False, False, False, False).Value

32        SaveFxDataToDictionary CCys, Numeraire, MarketWb, FxVolHandling, DCT

33        SaveCorrelationsToDictionary CCys, InflationIndices, Numeraire, MarketWb, DCT

34        For Each Ccy In CCys
35            If Not IsInCollection(MarketWb.Worksheets, UCase(Ccy)) Then Throw "The market data workbook must have a sheet called '" + UCase(Ccy) + "' with market data for " + UCase(Ccy), True
36            SaveCurrencySheetToDictionary MarketWb.Worksheets(UCase(CStr(Ccy))), CStr(Ccy), DCT
37        Next

          'Even if there are no credit curves to save, we still save funding spreads
38        SaveCreditDataToDictionary MarketWb, DCT, Credits

39        If Not (IsEmpty(InflationIndices) Or IsMissing(InflationIndices) Or IsNull(InflationIndices)) Then
40            For Each InflationIndex In InflationIndices
41                SaveInflationSheetToDictionary MarketWb.Worksheets(CStr(InflationIndex)), DCT
42            Next
43        End If

44        DataToWrite = ConvertToJson(DCT, 3, AS_RowByRow)

45        ThrowIfError sFileSave(FileName, DataToWrite, "")
46        SaveDataFromMarketWorkbookToFile = FileName

47        If OldSaved Then
48            If Not ThisWorkbook.Saved Then
49                ThisWorkbook.Saved = True
50            End If
51        End If

52        Exit Function
ErrHandler:
53        SaveDataFromMarketWorkbookToFile = "#SaveDataFromMarketWorkbookToFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveCurrencySheetToDictionary
' Author     : Philip Swannell
' Date       : 24-Jan-2018
' Purpose    : Saves the data on a currency sheet of this workbook to a dictionary, for later translation to JSON format
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveCurrencySheetToDictionary(ws As Worksheet, CurrencyCode As String, DCT As Dictionary)
          Dim i As Long
          Dim SwapRatesDict As New Dictionary
          Dim SwapsData As Variant
          Dim SwapsRange As Range
          Dim VolDataRange As Range
          Dim VolDict As New Dictionary
          Dim VolParameters As Variant
          Dim XccyBasisSpreadsData As Variant
          Dim XccyBasisSpreadsDict As New Dictionary
          Dim XccyBasisSpreadsRange As Range
          Dim FloatingLegType As String 'For Interest rate swaps, read the type (RFR or IBOR) from the sheet
          'For the time being assume all XCCY basis swaps will be RFR vs RFR. Is that true, even for EUR?
          Const FloatingLegType1 As String = "RFR"
          Const FloatingLegType2 As String = "RFR"

1         On Error GoTo ErrHandler

2         ws.Calculate

3         CheckName ws, "Title", 1, 1
4         CheckName ws, "SwapRatesInit", 1, 6
5         CheckName ws, "XccyBasisSpreadsInit", 1, 6
6         CheckName ws, "VolInit", 1, 1
7         CheckName ws, "SwaptionVolParameters", , 2
8         CheckName ws, "FloatingLegType", 1, 1

9         If Left(CStr(RangeFromSheet(ws, "Title")), 3) <> CurrencyCode Then
10            Throw "Assertion Failed - mismatch between CurrencyCode (" + CurrencyCode + ") and contents of the cell named 'Title' on sheet '" + _
                  ws.Name + "' of workbook '" + ws.Parent.Name + "' which is '" + sArrayMakeText(RangeFromSheet(ws, "Title"))(1, 1) + "'"
11        End If

12        Set SwapsRange = sExpandDown(RangeFromSheet(ws, "SwapRatesInit"))
13        Set XccyBasisSpreadsRange = sExpandDown(RangeFromSheet(ws, "XccyBasisSpreadsInit"))
14        FloatingLegType = ws.Range("FloatingLegType").Value

15        CheckSwapsData CurrencyCode, SwapsRange.Value2, SwapsData, SwapsRange.Rows(0).Value, FloatingLegType
16        For i = 1 To sNCols(SwapsData)
17            SwapRatesDict.Add SwapsData(1, i), To1D(sSubArray(SwapsData, 2, i, , 1))
18        Next i
19        DCT.Add "SwapRates_" & CurrencyCode, SwapRatesDict

20        CheckXccyBasisSpreadsData CurrencyCode, XccyBasisSpreadsRange.Value2, XccyBasisSpreadsData, RangeFromSheet(ws, "Spread_is_on", False, True, False, False, False), FloatingLegType1, FloatingLegType2
21        For i = 1 To sNCols(XccyBasisSpreadsData)
22            XccyBasisSpreadsDict.Add XccyBasisSpreadsData(1, i), To1D(sSubArray(XccyBasisSpreadsData, 2, i, , 1))
23        Next i
24        DCT.Add "XccyBasisSpreads_" & CurrencyCode, XccyBasisSpreadsDict

25        Set VolDataRange = Range(RangeFromSheet(ws, "VolInit"), RangeFromSheet(ws, "VolInit").End(xlDown).End(xlToRight))
26        With VolDataRange
27            CheckVolData CurrencyCode, .Offset(-1, -1).Resize(.Rows.Count + 1, .Columns.Count + 1).Value2
28        End With
29        VolParameters = RangeFromSheet(ws, "SwaptionVolParameters").Value
30        CheckVolParameters CurrencyCode, VolParameters

31        VolDict.Add "FixedFrequency", sParseFrequencyString(sVLookup("FixedFrequency", VolParameters), True, True)
32        VolDict.Add "FloatingFrequency", sParseFrequencyString(sVLookup("FloatingFrequency", VolParameters), True, True)
33        VolDict.Add "FixedDCT", sParseDCT(sVLookup("FixedDCT", VolParameters), True, True)
34        VolDict.Add "FloatingDCT", sParseDCT(sVLookup("FloatingDCT", VolParameters), True, True)
35        VolDict.Add "xlabels", To1D(VolDataRange.Columns(0).Value)
36        VolDict.Add "ylabels", To1D(VolDataRange.Rows(0).Value)
37        VolDict.Add "data", VolDataRange.Value
38        DCT.Add "SwaptionVols_" & CurrencyCode, VolDict
39        Exit Sub
40        Exit Sub
ErrHandler:
41        Throw "#SaveCurrencySheetToDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDtoTwoD
' Author     : Philip Swannell
' Date       : 08-Nov-2019
' Purpose    : Converts 1-dimensional array to two-dimensional array with 1 column and 1-based.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OneDtoTwoD(x)
          Dim i As Long
          Dim Res() As Variant
1         ReDim Res(1 To UBound(x) - LBound(x) + 1, 1 To 1)
2         For i = LBound(x) To UBound(x)
3             Res(i - LBound(x) + 1, 1) = x(i)
4         Next i
5         OneDtoTwoD = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveCorrelationsToDictionary
' Author    : Philip Swannell
' Date      : 23-Mar-2016
' Purpose   : Gets a correlation matrix from the appropriate sheet of a market data workbook
'              and saves it to a dictionary for subsequent streaming to JSON
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveCorrelationsToDictionary(ByVal RequiredCcys, ByVal RequiredInflation As Variant, Numeraire As String, MarketBook As Workbook, DCT As Dictionary)
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim N As Long
          Dim R As Range
          Dim RequiredLabels
          Dim rValue As Variant
          Dim SheetName As String
          Dim SmallCorrMatrix As Variant
          Dim VolDict As New Dictionary
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         If Numeraire <> RequiredCcys(1, 1) Then Throw "Assertion failed: Numeraire currency must be the first currency listed in CurrenciesRequired"
3         If Not (IsEmpty(RequiredInflation) Or IsNull(RequiredInflation)) Then
4             If NumDimensions(RequiredInflation) = 1 Then
5                 RequiredInflation = OneDtoTwoD(RequiredInflation)
6             End If
7             RequiredCcys = sArrayStack(RequiredCcys, RequiredInflation)
8         End If

9         SheetName = "HistoricalCorr" + Numeraire
10        If Not IsInCollection(MarketBook.Worksheets, SheetName) Then Throw "Cannot find sheet '" + SheetName + "' in workbook " + MarketBook.Name
11        Set ws = MarketBook.Worksheets(SheetName)
12        ws.Calculate
13        If Not IsInCollection(ws.Names, "HistCorrMatrix") Then Throw "Connot find named range HistCorrMatrix on sheet HistoricalCorr of market data workbook"
14        Set R = sExpandRightDown(RangeFromSheet(ws, "HistCorrMatrix"))
15        rValue = R.Value
16        If Not sArraysNearlyIdentical(rValue, sArrayTranspose(rValue)) Then Throw "Range HistCorrMatrix on sheet " + SheetName + " of the market data workbook contains data that's not symmetric"

17        RequiredLabels = sArrayConcatenate(RequiredCcys, " IR")
          'In the VBA layer we treat currencies and Inflation indices as different things, but in the R code there is a sense in which each inflation index is "just another currency"
18        If sNRows(RequiredCcys) > 1 Then
19            RequiredLabels = sArrayStack(RequiredLabels, sArrayConcatenate(sSubArray(RequiredCcys, 2), " FX"))
20        End If
21        N = sNRows(RequiredLabels)
22        MatchIDs = sMatch(RequiredLabels, R.Columns(1).Value)
23        Force2DArrayRMulti MatchIDs, RequiredLabels

24        For i = 1 To sNRows(MatchIDs)
25            If Not IsNumber(MatchIDs(i, 1)) Then Throw "Cannot find row labelled " + RequiredLabels(i, 1) + " in the range HistCorrMatrix on the sheet " + SheetName + " of the market data workbook"
26        Next i
27        ReDim SmallCorrMatrix(1 To N, 1 To N) As Variant

          'Construct SmallCorrMatrix
28        For i = 1 To N
29            For j = 1 To i
30                SmallCorrMatrix(i, j) = rValue(MatchIDs(i, 1), MatchIDs(j, 1))
31                SmallCorrMatrix(j, i) = rValue(MatchIDs(i, 1), MatchIDs(j, 1))
32            Next j
33        Next i

34        VolDict.Add "rownames", To1D(RequiredLabels)
35        VolDict.Add "colnames", To1D(RequiredLabels)
36        VolDict.Add "data", SmallCorrMatrix
37        DCT.Add "Correlations_" & Numeraire, VolDict    'append the numeraire so that files could save more than one correlation matrix - _
                                                           not needed for building a model but would be needed if we want to extend the file format _
                                                           to use for persistence of the data in a market data workbook

38        Exit Sub
ErrHandler:
39        Throw "#SaveCorrelationsToDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveCreditDataToDictionary
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Saves information from the Credit sheet of a market data workbook to a Dictionary for subsequent streaming to a JSON file
' Parameters :
'  MarketWb: A workbook such as this one.
'  DCT     : A (pre-existing) dictionary to which data is added
'  Credits : column array of credits, must be a subset of the leftmost column of the data held on the "Credit" sheet
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveCreditDataToDictionary(MarketWb As Workbook, DCT As Dictionary, Credits As Variant)

          Dim BlankLevels() As Double
          Dim BlankTenures() As String
          Dim CurveDct As Dictionary
          Dim EntireRange As Range
          Dim i As Long
          Dim j As Long
          Dim MatchIDs As Variant
          Dim NbCounterparties As Long
          Dim NbTenures
          Dim NumGood As Long
          Dim shCDS As Worksheet
          Dim SourceCDSLevels
          Dim SourceCounterparties
          Dim SourceFundingSpreads
          Dim SourceRecoveries
          Dim SourceTenures
          Dim ThisCDSLevels
          Dim ThisCounterParty As String
          Dim ThisFundingSpread
          Dim ThisRecovery
          Dim ThisRowNum As Long
          Dim ThisTenures

1         On Error GoTo ErrHandler
2         If Not IsInCollection(MarketWb.Worksheets, "Credit") Then Throw "Cannot find sheet 'Credit' in workbook " + MarketWb.Name
3         Set shCDS = MarketWb.Worksheets("Credit")
4         shCDS.Calculate

5         Set EntireRange = CDSRange(shCDS)
6         With EntireRange
7             Set SourceTenures = .Cells(1, 4).Resize(1, .Columns.Count - 3)
8             Set SourceCounterparties = .Cells(2, 1).Resize(.Rows.Count - 1)
9             Set SourceFundingSpreads = .Cells(2, 2).Resize(.Rows.Count - 1)
10            Set SourceRecoveries = .Cells(2, 3).Resize(.Rows.Count - 1)
11            Set SourceCDSLevels = .Cells(2, 4).Resize(.Rows.Count - 1, .Columns.Count - 3)
12        End With

          Dim FSdct As New Dictionary
          Dim MatchID
          Dim Names
          Dim Spreads
13        Names = SourceCounterparties.Value
14        MatchID = sMatch(gSELF, Names)
15        If Not IsNumber(MatchID) Then Throw ("Cannot find '" + gSELF + "' in list of credits on sheet " + shCredit.Name)
16        Spreads = SourceFundingSpreads.Value
17        If Not IsNumber(Spreads(MatchID, 1)) Then Spreads(MatchID, 1) = 0
18        For i = 1 To sNRows(Spreads)
19            If Not IsNumber(Spreads(i, 1)) Then Throw ("Found non number as funding spread for '" + CStr(Names(i, 1)) + "'")
20        Next i

21        FSdct.Add "Name", To1D(Names)
22        FSdct.Add "Spread", To1D(Spreads)
23        DCT.Add "FundingSpreads", FSdct

24        If IsEmpty(Credits) Or IsNull(Credits) Then Exit Sub

25        NbCounterparties = sNRows(Credits)
26        NbTenures = SourceTenures.Columns.Count
27        ReDim BlankTenures(1 To NbTenures)
28        ReDim BlankLevels(1 To NbTenures)
29        SourceTenures = SourceTenures.Value2
30        SourceCounterparties = SourceCounterparties.Value2
31        SourceCDSLevels = SourceCDSLevels.Value2
32        SourceFundingSpreads = SourceFundingSpreads.Value2
33        SourceRecoveries = SourceRecoveries.Value2
34        MatchIDs = sMatch(Credits, SourceCounterparties)
35        Force2DArray MatchIDs

36        For j = 1 To NbCounterparties
37            If Not IsNumber(MatchIDs(j, 1)) Then Throw "No credit data found for Counterparty '" + CStr(Credits(j, 1)) + "'"
38            ThisRowNum = MatchIDs(j, 1)

39            ThisCounterParty = SourceCounterparties(ThisRowNum, 1)
40            ThisFundingSpread = SourceFundingSpreads(ThisRowNum, 1)
41            ThisRecovery = SourceRecoveries(ThisRowNum, 1)
42            NumGood = 0
43            If Not IsNumber(ThisFundingSpread) Then
44                If ThisCounterParty = gSELF Then
45                    ThisFundingSpread = 0
46                Else
47                    Throw "Invalid funding spread for counterparty '" + ThisCounterParty + "'"
48                End If
49            End If

50            If Not IsNumber(ThisRecovery) Then Throw "Invalid recovery rate for counterparty '" + ThisCounterParty + "'"
51            If ThisRecovery < 0 Or ThisRecovery > 1 Then Throw "Invalid recovery rate for counterparty '" + ThisCounterParty + "'"

52            ThisCDSLevels = BlankLevels
53            ThisTenures = BlankTenures
54            For i = 1 To NbTenures
55                If IsNumber(SourceCDSLevels(ThisRowNum, i)) Then
56                    NumGood = NumGood + 1
57                    ThisTenures(NumGood) = SourceTenures(1, i)
58                    ThisCDSLevels(NumGood) = SourceCDSLevels(ThisRowNum, i)
59                End If
60            Next i
61            If NumGood = 0 Then Throw "No valid CDS levels found for counterparty '" + ThisCounterParty + "'"
62            If NumGood < NbTenures Then
63                ReDim Preserve ThisTenures(1 To NumGood)
64                ReDim Preserve ThisCDSLevels(1 To NumGood)
65            End If
66            Set CurveDct = New Dictionary
67            CurveDct.Add "party", ThisCounterParty
68            CurveDct.Add "recovery", ThisRecovery
69            CurveDct.Add "xlabels", ThisTenures
70            CurveDct.Add "data", ThisCDSLevels
71            DCT.Add "Credit_" + ThisCounterParty, CurveDct
72        Next j
73        Exit Sub
ErrHandler:
74        Throw "#SaveCreditDataToDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveFxDataToDictionary
' Author    : Philip Swannell
' Date      : 26-Jan-2018
' Purpose   : Reads the sheet FX and adds data to a dictionary for subsequent streaming
'             to a JSON file which will be picked up by R code.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SaveFxDataToDictionary(TheCcys As Variant, Numeraire As String, MarketWb As Workbook, FxVolHandling, ByRef DCT As Dictionary)

          Const SheetName = "FX"
          Dim CurveDct As Dictionary
          Dim DataDescription As String
          Dim DataDescription2 As String
          Dim FxVolHistoricLabel As String
          Dim i As Long
          Dim j As Long
          Dim LeftCol
          Dim MatchIDs As Variant
          Dim MatchIDs2 As Variant
          Dim NumPoints As Long
          Dim NumPoints2 As Long
          Dim RawData As Variant
          Dim RawData2 As Variant
          Dim RawDataRange As Range
          Dim RawDataRange2 As Range
          Dim RequiredPairs As Variant
          Dim RequiredPairsObverse As Variant
          Dim RowNumber As Long
          Dim shFx As Worksheet
          Dim Tmp As Variant
          Dim TopRow As Variant
          Dim WriteFxVol As Boolean
          Dim WriteFxVolHistoric As Boolean

1         On Error GoTo ErrHandler

2         If VarType(FxVolHandling) = vbBoolean Then
3             If FxVolHandling Then
4                 FxVolHandling = 1
5             Else
6                 FxVolHandling = 2
7             End If
8         End If

9         Select Case FxVolHandling
              Case 1 'Write both FxVol and FxVolHistoric. R code for building model could cope with this (special handling for Cayley project)
10                WriteFxVol = True
11                WriteFxVolHistoric = True
12                FxVolHistoricLabel = "FxVolHistoric_"
13            Case 2 'Write FxVol only
14                WriteFxVol = True
15                WriteFxVolHistoric = False
16                FxVolHistoricLabel = ""
17            Case 3 'Write FxVolHistoric, but label as FXVol. Required with Julia code as don't want to special case the Julia XVA code for this strange requirement of Cayley.
18                WriteFxVol = False
19                WriteFxVolHistoric = True
20                FxVolHistoricLabel = "FxVol_"
21            Case Else
22                Throw "Argument 'FXVolHandling' not recognised"
23        End Select

24        If Not TheCcys(1, 1) = Numeraire Then Throw "Assertion failed: first element of TheCCys must be the numeraire (which is currently " + Numeraire + ")"
25        If Not IsInCollection(MarketWb.Worksheets, SheetName) Then Throw "Cannot find sheet " + SheetName + " in market data workbook"
26        Set shFx = MarketWb.Worksheets(SheetName)
27        shFx.Calculate

28        Set RawDataRange = sExpandRightDown(RangeFromSheet(shFx, "FxDataTopLeft"))
29        NumPoints = RawDataRange.Columns.Count - 2
30        DataDescription = "range " + Replace(RawDataRange.Address, "$", "") + " of sheet '" + RawDataRange.Parent.Name + "' of workbook '" + RawDataRange.Parent.Parent.Name + "'"

31        If Not RawDataRange.Cells(1, 2).Value = "Spot" Then Throw "Unexpected error. Label 'Spot' not found as header for first column " + DataDescription
32        RawData = RawDataRange.Value
33        TopRow = RawDataRange.Rows(1).Value
34        If Not sRowAnd(sIsRegMatch("^[0-9]+(W|M|Y|D)$", sSubArray(TopRow, 1, 3)))(1, 1) Then Throw "Invalid label '" + CStr(Tmp) + "' found in top row of " + DataDescription

35        RequiredPairs = sArrayConcatenate(TheCcys, Numeraire)
36        RequiredPairsObverse = sArrayConcatenate(Numeraire, TheCcys)
37        LeftCol = RawDataRange.Columns(1).Value

38        MatchIDs = sMatch(RequiredPairs, LeftCol)
39        MatchIDs2 = sMatch(RequiredPairsObverse, LeftCol)
40        Force2DArrayRMulti MatchIDs, MatchIDs2

          Dim Spots() As Double
41        ReDim Spots(1 To sNRows(TheCcys))
42        Spots(1) = 1

43        For i = 2 To sNRows(RequiredPairs)
44            If IsNumber(MatchIDs(i, 1)) Then
45                Tmp = RawData(MatchIDs(i, 1), 2)
46                If Not IsNumber(Tmp) Then Throw "Invalid (non numeric) Fx rate found in second column of " + DataDescription
47                If Tmp <= 0 Then Throw "Invalid (non positive) Fx rate found in second column of " + DataDescription
48                Spots(i) = Tmp
49            ElseIf IsNumber(MatchIDs2(i, 1)) Then
50                Tmp = RawData(MatchIDs2(i, 1), 2)
51                If Not IsNumber(Tmp) Then Throw "Invalid (non numeric) Fx rate found in second column of " + DataDescription
52                If Tmp <= 0 Then Throw "Invalid (non positive) Fx rate found in second column of " + DataDescription
53                Spots(i) = 1 / Tmp
54            Else
55                Throw "Cannot find Fx Vol data for either " + RequiredPairs(i, 1) + " or " + RequiredPairsObverse(i, 1) + " in first column of " + DataDescription
56            End If
57        Next i

58        Set CurveDct = New Dictionary
59        CurveDct.Add "symbol", To1D(RequiredPairs)
60        CurveDct.Add "rate", Spots
61        DCT.Add "FxSpot", CurveDct

          Dim Data() As Double
          Dim Data2() As Double
          Dim xlabels
          Dim xlabelsHistoric

62        ReDim Data(1 To NumPoints)

63        If WriteFxVol Then
64            xlabels = To1D(sSubArray(TopRow, , 3))
65            For i = 2 To sNRows(RequiredPairs)
66                If IsNumber(MatchIDs(i, 1)) Then
67                    RowNumber = (MatchIDs(i, 1))
68                ElseIf IsNumber(MatchIDs2(i, 1)) Then
69                    RowNumber = (MatchIDs2(i, 1))
70                Else
71                    Throw "Cannot find Fx Vol data for either " + RequiredPairs(i, 1) + " or " + RequiredPairsObverse(i, 1) + " in the data held on sheet '" + shFx.Name + "' of the market data workbook"
72                End If
73                For j = 1 To NumPoints
74                    Tmp = RawData(RowNumber, j + 2)
75                    If Not (IsNumber(Tmp)) Then Throw "Non-numeric fx vol data found in " + DataDescription + " for currency pair " + RequiredPairs(i, 1)
76                    If Tmp < 0 Then Throw "Invalid (not positive) fx vol data found in " + DataDescription + " for currency pair " + RequiredPairs(i, 1)
77                    Data(j) = Tmp
78                Next j
79                Set CurveDct = New Dictionary
80                CurveDct.Add "xlabels", xlabels
81                CurveDct.Add "data", Data
82                DCT.Add "FxVol_" & RequiredPairs(i, 1), CurveDct
83            Next i
84        End If

85        If WriteFxVolHistoric Then
              'Although we only show 1 column on the sheet for the historical vols we "stretch" to the same number of columns as the market vols
86            Set RawDataRange2 = sExpandRightDown(RangeFromSheet(shFx, "HistoricFxVolsTopLeft"))
87            With RawDataRange2
88                xlabelsHistoric = To1D(.Rows(0).Offset(, 1).Resize(, .Columns.Count - 1))
89            End With

90            NumPoints2 = RawDataRange2.Columns.Count - 1
91            ReDim Data2(1 To NumPoints2)
92            RawData2 = RawDataRange2.Value2

93            DataDescription2 = "range " + Replace(RawDataRange2.Address, "$", "") + " of sheet '" + RawDataRange2.Parent.Name + "' of workbook '" + RawDataRange2.Parent.Parent.Name + "'"
94            LeftCol = RawDataRange2.Columns(1).Value

95            MatchIDs = sMatch(RequiredPairs, LeftCol)
96            MatchIDs2 = sMatch(RequiredPairsObverse, LeftCol)
97            Force2DArrayRMulti MatchIDs, MatchIDs2

98            For i = 2 To sNRows(RequiredPairs)
99                If IsNumber(MatchIDs(i, 1)) Then
100                   RowNumber = (MatchIDs(i, 1))
101               ElseIf IsNumber(MatchIDs2(i, 1)) Then
102                   RowNumber = (MatchIDs2(i, 1))
103               Else
104                   Throw "Cannot find historical Fx Vol data for either " + RequiredPairs(i, 1) + " or " + RequiredPairsObverse(i, 1) + " in the data held on sheet '" + shFx.Name + "' of the market data workbook"
105               End If
                    
106               For j = 1 To NumPoints2
107                   Tmp = RawData2(RowNumber, j + 1)
108                   If Not (IsNumber(Tmp)) Then Throw "Non-numeric historic fx vol data found in " + DataDescription2 + " for currency pair " + RequiredPairs(i, 1)
109                   If Tmp < 0 Then Throw "Invalid (not positive) historic fx vol data found in " + DataDescription + " for currency pair " + RequiredPairs(i, 1)
110                   Data2(j) = Tmp
111               Next j
112               Set CurveDct = New Dictionary
113               CurveDct.Add "xlabels", xlabelsHistoric
114               CurveDct.Add "data", Data2
115               DCT.Add FxVolHistoricLabel & RequiredPairs(i, 1), CurveDct
116           Next i
117       End If

118       Exit Function
ErrHandler:
119       Throw "#SaveFxDataToDictionary (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveInflationSheetToDictionary
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    :
' Parameters :
'  ws : A worksheet of a market data workbook designed to store data for an inflation index
'  DCT:  A pre-existing Dictionary to which this method adds more items
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveInflationSheetToDictionary(ws As Worksheet, DCT As Dictionary)
1         On Error GoTo ErrHandler
          Dim CurveDict As New Dictionary
          Dim HistoricSetsDict As New Dictionary
          Dim i As Long
          Dim ParamsDict As New Dictionary
          Dim VolDict As New Dictionary

          Dim AnchorDate As Long
          Dim EffectiveDateMethod As String
          Dim HistoricSets()
          Dim IndexName As String
          Dim InflationParameters
          Dim LagMethod As String
          Dim SeasonalAdjustments As Variant
          Dim VolData As Variant
          Dim VolRange As Range
          Dim ZCSwapsData As Variant
          Dim ZCSwapsRange As Range

2         ws.Calculate
3         IndexName = ws.Name
4         CheckName ws, "ZCSwapsInit", 1, 2
5         CheckName ws, "SeasonalAdjustments", 12, 2
6         CheckName ws, "HistoricDataInit", 1, 3
7         CheckName ws, "ParametersInit", 1, 2
8         CheckName ws, "Parameters", , 2
9         CheckName ws, "InflationVolInit", 1, 2

10        CheckInflationParameters ws.Range("Parameters"), InflationParameters
11        For i = 1 To sNRows(InflationParameters)
12            ParamsDict.Add InflationParameters(i, 1), InflationParameters(i, 2)
13        Next i
14        CurveDict.Add "Parameters", ParamsDict
15        EffectiveDateMethod = ParamsDict("EffectiveDate")
16        LagMethod = ParamsDict("LagMethod")
17        AnchorDate = RangeFromSheet(shConfig, "AnchorDate")

18        Set ZCSwapsRange = sExpandDown(ws.Range("ZCSwapsInit"))
19        CheckZCInflationSwapsData IndexName, ZCSwapsRange.Value, ZCSwapsData
20        CurveDict.Add "xlabels", To1D(ZCSwapsRange.Columns(1).Value)
21        CurveDict.Add "data", To1D(ZCSwapsRange.Columns(2).Value)
22        CheckHistoricInflation sExpandDown(ws.Range("HistoricDataInit")), HistoricSets, AnchorDate    ', EffectiveDateMethod, IndexName
23        For i = 2 To sNRows(HistoricSets)
24            HistoricSets(i, 1) = Format(HistoricSets(i, 1), "yyyy-mm-dd")
25        Next i

26        For i = 1 To sNCols(HistoricSets)
27            HistoricSetsDict.Add HistoricSets(1, i), To1D(sSubArray(HistoricSets, 2, i, , 1))
28        Next i
29        CurveDict.Add "HistoricSets", HistoricSetsDict
30        CheckSeasonalAdjustments RangeFromSheet(ws, "SeasonalAdjustments"), SeasonalAdjustments
          'PGS 24-May-2017 R code now handles seasonal adjustment just as they appear on the sheets - vector of 12 numbers summing to zero
31        CurveDict.Add "SeasonalAdjustments", To1D(SeasonalAdjustments)
32        Set VolRange = sExpandDown(ws.Range("InflationVolInit"))
33        CheckInflationVolData IndexName, VolRange.Value, VolData
34        VolDict.Add "xlabels", To1D(VolRange.Columns(1).Value)
35        VolDict.Add "data", To1D(VolRange.Columns(2).Value)
36        DCT.Add "Inflation_" & IndexName, CurveDict
37        DCT.Add "InflationVols_" & IndexName, VolDict
38        Exit Sub
ErrHandler:
39        Throw "#SaveInflationSheetToDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckInflationParameters
' Author    : Philip Swannell
' Date      : 27-Apr-2017
' Purpose   : Validator for inflation parameters
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckInflationParameters(R As Range, ByRef DataToSave As Variant)

          Dim BaseCurrency
          Dim EffectiveDate
          Dim ErrString
          Dim HistoricalVol
          Dim HistoricalVolSA
          Dim LagMethod

1         On Error GoTo ErrHandler
2         ErrString = "Error in range 'Parameters' on worksheet '" + R.Parent.Name + "' in workbook '" + R.Parent.Parent.Name + "' "

3         BaseCurrency = sVLookup("BaseCurrency", R.Value)
4         If CStr(BaseCurrency) = "#Not found!" Then
5             Throw ErrString + "Label 'BaseCurrency' not found in left column (i.e. in range " + Replace(R.Columns(1).Address, "$", "")
6         ElseIf Not IsNumber(sMatch(BaseCurrency, sCurrencies(False, True))) Then
7             Throw ErrString + "BaseCurrency '" + BaseCurrency + "' is not valid"
8         End If

9         LagMethod = sVLookup("LagMethod", R.Value)
10        If CStr(LagMethod) = "#Not found!" Then
11            Throw ErrString + "Label 'LagMethod' not found in left column (i.e. in range " + Replace(R.Columns(1).Address, "$", "")
12        ElseIf Not IsNumber(sMatch(LagMethod, sSupportedInflationLagMethods())) Then
13            Throw ErrString + "LagMethod '" + LagMethod + "' is not valid. Allowed values are :" + sTokeniseString(sSupportedInflationLagMethods(), ", ")
14        End If

15        EffectiveDate = sVLookup("EffectiveDate", R.Value)
16        If CStr(EffectiveDate) = "#Not found!" Then
17            Throw ErrString + "Label 'EffectiveDate' not found in left column (i.e. in range " + Replace(R.Columns(1).Address, "$", "")
18        ElseIf Not IsNumber(sMatch(EffectiveDate, sSupportedInflationEffectiveDates())) Then
19            Throw ErrString + "EffectiveDate '" + EffectiveDate + "' is not valid. Allowed values are :" + sTokeniseString(sSupportedInflationEffectiveDates(), ", ")
20        End If

21        HistoricalVol = sVLookup("HistoricalVol", R.Value)
22        If CStr(HistoricalVol) = "#Not found!" Then
23            Throw ErrString + "Label 'HistoricalVol' not found in left column (i.e. in range " + Replace(R.Columns(1).Address, "$", "")
24        ElseIf Not IsNumber(HistoricalVol) Then
25            Throw ErrString + "HistoricalVol must be a number"
26        End If

27        HistoricalVolSA = sVLookup("HistoricalVolSA", R.Value)
28        If CStr(HistoricalVolSA) = "#Not found!" Then
29            Throw ErrString + "Label 'HistoricalVolSA' not found in left column (i.e. in range " + Replace(R.Columns(1).Address, "$", "")
30        ElseIf Not IsNumber(HistoricalVolSA) Then
31            Throw ErrString + "HistoricalVolSA must be a number"
32        End If

33        DataToSave = R.Value

34        Exit Function
ErrHandler:
35        Throw "#CheckInflationParameters (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckHistoricInflation
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : Validation for data saved as Historic Sets on an inflation sheet
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckHistoricInflation(R As Range, ByRef DataToSave(), AnchorDate As Long)    ', EffectiveDateMethod As String, Index As String)

          Dim Data As Variant
          Dim ErrString As String
          Dim i As Long
          Dim j As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         ErrString = "Error in data for historic inflation sets on worksheet '" + R.Parent.Name + "' in workbook '" + R.Parent.Parent.Name + "'"

3         If R.Columns.Count <> 3 Then Throw ErrString + " data has " + CStr(R.Columns.Count) + " columns, but should have 3"
4         If R.Cells(0, 1).Value <> "Year" Then Throw ErrString + ". Header at cell " + Replace(R.Cells(0, 1).Address, "$", "") + " should read 'Year'"
5         If R.Cells(0, 2).Value <> "Month" Then Throw ErrString + ". Header at cell " + Replace(R.Cells(0, 1).Address, "$", "") + " should read 'Month'"
6         If R.Cells(0, 3).Value <> "Index" Then Throw ErrString + ". Header at cell " + Replace(R.Cells(0, 1).Address, "$", "") + " should read 'Index'"

7         Data = R.Value
8         NR = sNRows(Data)

          'Check for numbers in all three columns and whole numbers in left two columns
9         For i = 1 To NR
10            For j = 1 To 3
11                If Not IsNumberOrDate(Data(i, j)) Then Throw ErrString + ". Non-number at cell " + Replace(R.Cells(i, j).Address, "$", "")
12            Next j
13        Next i

14        If Data(1, 1) <> CLng(Data(1, 1)) Then Throw ErrString + ". Number at cell " + Replace(R.Cells(i, j).Address, "$", "") + " must be a whole number (to indicate year)"
15        If Data(1, 2) <> CLng(Data(1, 2)) Then Throw ErrString + ". Number at cell " + Replace(R.Cells(i, j).Address, "$", "") + " must be a whole number in the range 1 to 12 (to indicate month)"
16        If Data(1, 2) < 1 Or Data(1, 2) > 12 Then Throw ErrString + ". Number at cell " + Replace(R.Cells(i, j).Address, "$", "") + " must be a whole number in the range 1 to 12 (to indicate month)"

17        For i = 2 To NR
18            If Data(i, 1) <> Data(i - 1, 1) + IIf(Data(i - 1, 2) = 12, 1, 0) Then Throw ErrString + ". Out-of sequence year found at cell " + Replace(R.Cells(i, 1).Address, "$", "")
19            If Data(i, 2) <> Data(i - 1, 2) + IIf(Data(i - 1, 2) = 12, -11, 1) Then Throw ErrString + ". Out-of sequence month found at cell " + Replace(R.Cells(i, 2).Address, "$", "")
20        Next i

          'For US Release dates see:
          'https://www.bls.gov/schedule/news_release/cpi.htm
          'Between 14th and 18th of month?
          'For Euro Release dates see
          ' http://ec.europa.eu/eurostat/documents/272892/272971/HICP+Flash+estimate+release+calendar/a5b6c5bd-f3fe-4980-8b94-3433b689b26e
          'or from http://ec.europa.eu/eurostat/news/release-calendar
          'Between 16th and 22nd?
          'So Some time between 16 and 22 Feb, the index for _Jan_ is published...

          Dim LastMIs As Long
          Dim LastMShouldBe As Long
          Dim LastYIs As Long
          Dim LastYShouldBe As Long
21        LastYIs = Data(NR, 1): LastMIs = Data(NR, 2)
22        If Day(AnchorDate) <= 16 Then
              'The month before the previous month
23            LastMShouldBe = Month(DateSerial(Year(AnchorDate), Month(AnchorDate) - 2, 1))
24            LastYShouldBe = Year(DateSerial(Year(AnchorDate), Month(AnchorDate) - 2, 1))
25        Else
              'The previous month
26            LastMShouldBe = Month(DateSerial(Year(AnchorDate), Month(AnchorDate) - 1, 1))
27            LastYShouldBe = Year(DateSerial(Year(AnchorDate), Month(AnchorDate) - 1, 1))
28        End If

          Dim MonthDiff As Long    ' If MonthDiff is positive then we have too many dates, if negative then we have too few

29        MonthDiff = 12 * LastYIs + LastMIs - (12 * LastYShouldBe + LastMShouldBe)

30        Select Case MonthDiff
              Case Is >= 1
31                If (MonthDiff = 1 And Day(AnchorDate) < 10) Or MonthDiff > 1 Then
32                    Throw "There appear to be too many sets entered in the Historic sets range of the worksheet '" + R.Parent.Name + "' of the market data workbook ('" + R.Parent.Parent.Name + "'). The last row should be for " + Format(DateSerial(LastYShouldBe, LastMShouldBe, 1), "mmm yyyy") + " but instead it's for " + Format(DateSerial(LastYIs, LastMIs, 1), "mmm yyyy")
33                End If
34            Case Is <= -1
35                If (MonthDiff = -1 And Day(AnchorDate) > 23) Or MonthDiff < -1 Then
36                    Throw "There appear to be too few sets entered in the Historic sets range of the worksheet '" + R.Parent.Name + "' of the market data workbook ('" + R.Parent.Parent.Name + "'). The last row should be for " + Format(DateSerial(LastYShouldBe, LastMShouldBe, 1), "mmm yyyy") + " but instead it's for " + Format(DateSerial(LastYIs, LastMIs, 1), "mmm yyyy")
37                End If
38        End Select

          Dim RDotV As Variant
39        RDotV = R.Value

40        ReDim DataToSave(1 To R.Rows.Count + 1, 1 To 2)
41        DataToSave(1, 1) = "FirstOfMonth": DataToSave(1, 2) = "Index"
42        For i = 1 To R.Rows.Count
43            DataToSave(i + 1, 1) = CLng(DateSerial(RDotV(i, 1), RDotV(i, 2), 1))
44            DataToSave(i + 1, 2) = RDotV(i, 3)
45        Next i

46        Exit Function
ErrHandler:
47        Throw "#CheckHistoricInflation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckSeasonalAdjustments
' Author    : Philip Swannell
' Date      : 03-May-2017
' Purpose   : Validator for the seasonal adjustments range
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckSeasonalAdjustments(R As Range, ByRef SeasonalAdjustments As Variant)

          Dim c As Range
          Dim SheetFullName As String
          Dim Sum As Double

1         On Error GoTo ErrHandler
2         SheetFullName = "worksheet '" + R.Parent.Name + "' in workbook '" + R.Parent.Parent.Name + "'"

3         If R.Rows.Count <> 12 Then Throw "Range 'SeasonalAdjustments' on " + SheetFullName + " has " + CStr(R.Rows.Count) + " rows but should have 12"
4         If R.Columns.Count <> 2 Then Throw "Range 'SeasonalAdjustments' on " + SheetFullName + " has " + CStr(R.Columns.Count) + " columns but should have 2"
5         Sum = 0
6         For Each c In R.Columns(2).Cells
7             If Not IsNumberOrDate(c.Value2) Then
8                 Throw "Non number found (at cell " + Replace(c.Address, "$", "") + ") in the range 'SeasonalAdjustments' on " + SheetFullName
9             End If
10            Sum = Sum + c.Value2
11        Next c

12        If Abs(Sum) > 0.000000001 Then
13            Throw "12 numbers in the right column of the range 'SeasonalAdjustments' (" + _
                  Replace(R.Columns(2).Address, "$", "") + ") on " + SheetFullName + " should sum to zero, but their sum is " + CStr(Sum)
14        End If
15        SeasonalAdjustments = R.Columns(2).Value2

16        Exit Function
ErrHandler:
17        Throw "#CheckSeasonalAdjustments (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckName
' Author    : Philip Swannell
' Date      : 24-Mar-2016
' Purpose   : To allow for simple tests as to whether the currency sheet is mal-formed
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckName(ws As Worksheet, Name As String, Optional NumRowsShouldBe As Long, Optional NumColumnsShouldBe As Long)
          Dim ErrString As String
          Dim N As Name
          Dim NameExists As Boolean
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set N = ws.Names(Name)
3         NameExists = True
4         Set R = N.RefersToRange
5         If Not R.Parent Is ws Then
6             ErrString = "Name '" + Name + "' on workSheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "' refers to a range on a different sheet"
7             GoTo ErrHandler
8         End If
9         If NumRowsShouldBe <> 0 Then
10            If R.Rows.Count <> NumRowsShouldBe Then
11                ErrString = "Name '" + Name + "' on workSheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "' should refer to a range with " + CStr(NumRowsShouldBe) + " rows, but it refers to " + CStr(Replace(R.Address, "$", "")) + " which has " + CStr(R.Rows.Count)
12                GoTo ErrHandler
13            End If
14        End If
15        If NumColumnsShouldBe <> 0 Then
16            If R.Columns.Count <> NumColumnsShouldBe Then
17                ErrString = "Name '" + Name + "' on workSheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "' should refer to a range with " + CStr(NumColumnsShouldBe) + " columns, but it refers to " + CStr(Replace(R.Address, "$", "")) + " which has " + CStr(R.Columns.Count)
18                GoTo ErrHandler
19            End If
20        End If

21        Exit Function
ErrHandler:
22        If ErrString = "" Then
23            If NameExists Then
24                ErrString = "Name '" + Name + " on workSheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "' does not refer to a range"
25            Else
26                ErrString = "WorkSheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "' does not have a named range '" + Name + "'"
27            End If
28        End If
29        Throw "#CheckName (line " & CStr(Erl) + "): " & ErrString & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckVolParameters
' Author    : Philip Swannell
' Date      : 12-May-2016
' Purpose   : Validation for the SwaptionVolParameters that should appear on the currency
'             sheets giving information about the quoting conventions for the swaps whose vol is given
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckVolParameters(CurrencyCode As String, VolParameters)
          Dim FixedDCT As String
          Dim FixedFrequency As String
          Dim FloatingDCT As String
          Dim FloatingFrequency As String
          Dim Res
1         On Error GoTo ErrHandler
2         If sNRows(sCompareTwoArrays(sTokeniseString("FixedFrequency,FloatingFrequency,FixedDCT,FloatingDCT"), sSubArray(VolParameters, 1, 1, , 1), "In1AndNotIn2")) > 1 Then
3             Throw "Left column of range SwaptionVolParameters for currency " + CurrencyCode + " must contain the labels FixedFrequency, FloatingFrequency, FixedDCT, FloatingDCT"
4         End If

5         FixedFrequency = sVLookup("FixedFrequency", VolParameters)
6         FloatingFrequency = sVLookup("FloatingFrequency", VolParameters)
7         FixedDCT = sVLookup("FixedDCT", VolParameters)
8         FloatingDCT = sVLookup("FloatingDCT", VolParameters)

9         Res = sParseFrequencyString(FixedFrequency, False)
10        If Not IsNumber(Res) Then Throw "Invalid FixedFrequency found in the SwaptionVolParameters range for currency " + CurrencyCode + " " + Res
11        Res = sParseFrequencyString(FloatingFrequency, False)
12        If Not IsNumber(Res) Then Throw "Invalid FloatingFrequency found in the SwaptionVolParameters range for currency " + CurrencyCode + " " + Res
13        Res = sParseDCT(FixedDCT, False, False)
14        If Left(Res, 1) = "#" Then Throw "Invalid FixedDCT found in the SwaptionVolParameters range for currency " + CurrencyCode + " " + Res
15        Res = sParseDCT(FloatingDCT, True, False)
16        If Left(Res, 1) = "#" Then Throw "Invalid FloatingDCT found in the SwaptionVolParameters range for currency " + CurrencyCode + " " + Res
17        Exit Function
ErrHandler:
18        Throw "#CheckVolParameters (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckVolData
' Author    : Philip Swannell
' Date      : 24-Mar-2016
' Purpose   : Validation
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckVolData(CurrencyCode As String, DataToCheck As Variant)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         NR = sNRows(DataToCheck): NC = sNCols(DataToCheck)

3         For i = 2 To NR
4             For j = 2 To NC
5                 If Not IsNumber(DataToCheck(i, j)) Then Throw "Found non-number in volatility data for " + CurrencyCode
6             Next
7         Next

8         For i = 2 To NR
9             If Not CheckTenureString(DataToCheck(i, 1)) Then Throw "Bad tenure string in volatility data for " + CurrencyCode
10        Next

11        For j = 2 To NC
12            If Not CheckTenureString(DataToCheck(1, j)) Then Throw "Bad tenure string in volatility data for " + CurrencyCode
13        Next

14        Exit Function
ErrHandler:
15        Throw "#CheckVolData (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckTenureString
' Author    : Philip Swannell
' Date      : 24-Mar-2016
' Purpose   : Yet more validation
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckTenureString(TenureString As Variant) As Boolean

1         On Error GoTo ErrHandler
2         If VarType(TenureString) <> vbString Then Exit Function
3         Select Case Right(TenureString, 1)
              Case "M", "Y", "W", "D"
4             Case Else
5                 Exit Function
6         End Select
7         If Not IsNumeric(Left(TenureString, Len(TenureString) - 1)) Then Exit Function
8         CheckTenureString = True
9         Exit Function
ErrHandler:
10        Throw "#CheckTenureString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckSwapsData
' Author    : Philip Swannell
' Date      : 27-Apr-2016
' Purpose   : Validation of the data held for swap rates on the Ccy_??? sheets of the market
'             data workbook
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CheckSwapsData(CurrencyCode, InputData, ByRef OutputData, InputHeaders, FloatingLegType As String)
          'It's safe to assume that InputData has 6 columns, since method CheckName will have been called
          Dim cnFixedDCT As Long
          Dim cnFixedFrequency As Long
          Dim cnFloatingDCT As Long
          Dim cnFloatingFrequency As Long
          Dim cnRates As Long
          Dim cnTenors As Long
          Dim i As Long
          Dim MatchIDsFixedDCT
          Dim MatchIDsFloatingDCT
          Dim SDCTS As Variant
          Dim Tmp As Variant
1         On Error GoTo ErrHandler
          Dim Headers(1 To 1, 1 To 7)

2         Select Case FloatingLegType
              Case "RFR", "IBOR"
3             Case Else
4                 Throw "Range FloatingLegType on worksheet " + CurrencyCode + " must be either 'IBOR' or 'RFR' but instead it is '" + FloatingLegType + "'."
5         End Select

6         For i = 1 To 6
7             Select Case InputHeaders(1, i)
                  Case "Tenor"
8                     Headers(1, i) = "Tenors"
9                     cnTenors = i
10                Case "Rate"
11                    Headers(1, i) = "Rates"
12                    cnRates = i
13                Case "FixFreq"
14                    Headers(1, i) = "FixedFrequency"
15                    cnFixedFrequency = i
16                Case "FloatFreq"
17                    Headers(1, i) = "FloatingFrequency"
18                    cnFloatingFrequency = i
19                Case "FixDCT"
20                    Headers(1, i) = "FixedDCT"
21                    cnFixedDCT = i
22                Case "FloatDCT"
23                    Headers(1, i) = "FloatingDCT"
24                    cnFloatingDCT = i
25                Case Else
26                    Throw "Unrecogised header name '" + CStr(InputHeaders(1, i)) + "' in swaps range for currency " + CStr(CurrencyCode)
27            End Select
28        Next
29        Headers(1, 7) = "FloatingLegType"

30        SDCTS = sSupportedDCTs()
31        MatchIDsFixedDCT = sMatch(sSubArray(InputData, 1, cnFixedDCT, , 1), SDCTS)
32        MatchIDsFloatingDCT = sMatch(sSubArray(InputData, 1, cnFloatingDCT, , 1), SDCTS)

33        OutputData = sArrayRange(InputData, sReshape(FloatingLegType, sNRows(InputData), 1))

34        For i = 1 To sNRows(InputData)
35            If Not CheckTenureString(InputData(i, cnTenors)) Then Throw "Invalid tenure string found at row " + CStr(i) + " column " + CStr(cnTenors) + " of the swaps data for " + CurrencyCode
36            If Not IsNumber(InputData(i, cnRates)) Then Throw "Invalid swap rate found at row " + CStr(i) + " column " + CStr(cnRates) + " of the swaps data for " + CurrencyCode
37            Tmp = sParseFrequencyString(CStr(InputData(i, cnFixedFrequency)), False)
38            If Not IsNumber(Tmp) Then
39                Throw "Invalid FixFreq found at row " + CStr(i) + " column " + CStr(cnFixedFrequency) + " of the swaps data for " + CurrencyCode + ": '" + Tmp + "'"
40            Else
41                OutputData(i, cnFixedFrequency) = Tmp
42            End If
43            Tmp = sParseFrequencyString(CStr(InputData(i, cnFloatingFrequency)), False)
44            If Not IsNumber(Tmp) Then
45                Throw "Invalid FloatFreq found at row " + CStr(i) + " column " + CStr(cnFloatingFrequency) + " of the swaps data for " + CurrencyCode + ": '" + Tmp + "'"
46            Else
47                OutputData(i, cnFloatingFrequency) = Tmp
48            End If
49            If Not IsNumber(MatchIDsFixedDCT(i, 1)) Then
50                Throw "Invalid fixed day count found at row " + CStr(i) + " column " + CStr(cnFixedDCT) + " of the swaps data for " + CurrencyCode + ". Supported day counts (case-sensitive) are: " + sConcatenateStrings(SDCTS, ", ")
51            ElseIf InputData(i, cnFixedDCT) <> SDCTS(MatchIDsFixedDCT(i, 1), 1) Then    ' unfortunately sMatch is case-insensitive, but we need case-sensitive checking
52                Throw "Invalid fixed day count found at row " + CStr(i) + " column " + CStr(cnFixedDCT) + " of the swaps data for " + CurrencyCode + ". Supported day counts (case-sensitive) are: " + sConcatenateStrings(SDCTS, ", ")
53            End If
54            If Not IsNumber(MatchIDsFloatingDCT(i, 1)) Then
55                Throw "Invalid floating day count found at row " + CStr(i) + " column " + CStr(cnFloatingDCT) + " of the swaps data for " + CurrencyCode + ". Supported day counts (case-sensitive) are: " + sConcatenateStrings(SDCTS, ", ")
56            ElseIf InputData(i, cnFloatingDCT) <> SDCTS(MatchIDsFloatingDCT(i, 1), 1) Then    ' unfortunately sMatch is case-insensitive, but we need case-sensitive checking
57                Throw "Invalid floating day count found at row " + CStr(i) + " column " + CStr(cnFloatingDCT) + " of the swaps data for " + CurrencyCode + ". Supported day counts (case-sensitive) are: " + sConcatenateStrings(SDCTS, ", ")
58            End If

59        Next i
60        OutputData = sArrayStack(Headers, OutputData)
61        Exit Sub
ErrHandler:
62        Throw "#CheckSwapsData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
'
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckZCInflationSwapsData
' Author    : Philip Swannell
' Date      : 25-Apr-2017
' Purpose   : validator
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckZCInflationSwapsData(InflationIndex As String, InputData, ByRef OutputData)
          Dim Headers(1 To 2, 1 To 2)
          Dim i As Long
1         On Error GoTo ErrHandler
2         Headers(1, 1) = "CHAR": Headers(1, 2) = "DOUBLE"
3         Headers(2, 1) = "xlabels": Headers(2, 2) = "data"    '<- list names required by R method addCurve.inflation
4         For i = 1 To sNRows(InputData)
5             If Not CheckTenureString(InputData(i, 1)) Then Throw "Invalid tenure string found at row " + CStr(i) + " column 1 of the zero coupon inflation swaps data for " + InflationIndex
6             If Not IsNumber(InputData(i, 2)) Then Throw "Invalid swap rate found at row " + CStr(i) + " column 2 of the zero coupon inflation swaps data for " + InflationIndex
7         Next i
8         OutputData = sArrayStack(Headers, InputData)
9         Exit Function
ErrHandler:
10        Throw "#CheckZCInflationSwapsData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckInflationVolData
' Author    : Philip
' Date      : 23-Jun-2017
' Purpose   : Validator
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckInflationVolData(InflationIndex As String, InputData, ByRef OutputData)
          Dim Headers(1 To 2, 1 To 2)
          Dim i As Long
1         On Error GoTo ErrHandler
2         Headers(1, 1) = "CHAR": Headers(1, 2) = "DOUBLE"
3         Headers(2, 1) = "xlabels": Headers(2, 2) = "data"
4         For i = 1 To sNRows(InputData)
5             If Not CheckTenureString(InputData(i, 1)) Then Throw "Invalid tenure string found at row " + CStr(i) + " column 1 of the inflation vol data for " + InflationIndex
6             If Not IsNumber(InputData(i, 2)) Then Throw "Invalid volatility found at row " + CStr(i) + " column 2 of the inflation vol data for " + InflationIndex
7         Next i
8         OutputData = sArrayStack(Headers, InputData)
9         Exit Function
ErrHandler:
10        Throw "#CheckInflationVolData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckXccyBasisSpreadsData
' Author    : Philip Swannell
' Date      : 27-Apr-2016
' Purpose   : Validation of the data held for cross currency basis spreads rates on the
'             currency sheets of the market data workbook
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CheckXccyBasisSpreadsData(CurrencyCode As String, InputData, ByRef OutputData, SpreadIsOn, FloatingLegType1 As String, FloatingLegType2 As String)
          'It's safe to assume that InputData has 6 columns, since method CheckName will have been called
          Dim CollateralCcy As String
          Dim ColNo As Long
          Dim i As Long
          Dim k As Long
          Dim NR As Long
          Dim SDCTS
          Dim Tmp As Variant
1         On Error GoTo ErrHandler
          Dim Headers(1 To 1, 1 To 9)
2         Headers(1, 1) = "Tenors": Headers(1, 2) = "Rates": Headers(1, 3) = "Freq1": Headers(1, 4) = "DCT1": Headers(1, 5) = "Freq2": Headers(1, 6) = "DCT2": Headers(1, 7) = "SpreadIsOn"
3         Headers(1, 8) = "FloatingLegType1": Headers(1, 9) = "FloatingLegType2"

4         SDCTS = sSortedArray(sSupportedDCTs())
5         CollateralCcy = RangeFromSheet(shConfig, "CollateralCcy", False, True, False, False, False).Value
6         If Not ((SpreadIsOn = CurrencyCode) Or (SpreadIsOn = CollateralCcy)) Then
7             Throw "Invalid 'SpreadIsOn' found for cross currency basis swap data for " + CStr(CurrencyCode) + _
                  ". This must be either '" + CurrencyCode + "' or '" + CollateralCcy + ",' but it is '" + CStr(SpreadIsOn) + "'. See cell " + _
                  Replace(RangeFromSheet(ThisWorkbook.Worksheets(CurrencyCode), "Spread_Is_On").Address, "$", "") + " of worksheet " + _
                  CurrencyCode + " of the market data workbook"
8         End If

9         NR = sNRows(InputData)
10        OutputData = sArrayRange(InputData, sReshape(SpreadIsOn, NR, 1), sReshape(FloatingLegType1, NR, 1), sReshape(FloatingLegType2, NR, 1))
11        For i = 1 To sNRows(InputData)
12            If Not CheckTenureString(InputData(i, 1)) Then Throw "Invalid tenure string found at row " + CStr(i) + " column 1 of the cross currency basis swaps data for " + CurrencyCode
13            If Not IsNumber(InputData(i, 2)) Then Throw "Invalid rate found at row " + CStr(i) + " column 2 of the cross currency basis swaps data for " + CurrencyCode

              'Freq
14            For k = 1 To 2
15                ColNo = Choose(k, 3, 5)
16                Tmp = sParseFrequencyString(CStr(InputData(i, ColNo)), False)
17                If Not IsNumber(Tmp) Then
18                    Throw "Invalid Freq" + CStr(k) + " found at row " + CStr(i) + " column " + CStr(ColNo) + " of the cross currency basis swaps data for " + CurrencyCode + " " + Tmp
19                Else
20                    OutputData(i, ColNo) = Tmp
21                End If
22            Next k
              'DCT
23            For k = 1 To 2
24                ColNo = Choose(k, 4, 6)
25                If Not IsNumber(sMatch(InputData(i, ColNo), SDCTS, True)) Then
26                    Throw "Invalid DCT" + CStr(k) + " found at row " + CStr(i) + " column " + CStr(ColNo) + " of the cross currency basis swaps data for " + CurrencyCode + " " + Tmp
27                End If
28            Next k

29        Next i
30        OutputData = sArrayStack(Headers, OutputData)
31        Exit Sub
ErrHandler:
32        Throw "#CheckXccyBasisSpreadsData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CDSRange
' Author    : Philip Swannell
' Date      : 16-Jan-2017
' Purpose   : Returns the range occupied by the cds data, including left headers and a
'             single row of top headers but excluding the CDS tickers
' -----------------------------------------------------------------------------------------------------------------------
Function CDSRange(shCDS As Worksheet) As Range
          Dim NumCols As Long
1         On Error GoTo ErrHandler
2         NumCols = RangeFromSheet(shCDS, "CDSTickersTopLeft").Column - RangeFromSheet(shCDS, "CDSTopLeft").Column
3         Set CDSRange = sExpandRightDown(RangeFromSheet(shCDS, "CDSTopLeft")).Resize(, NumCols)
4         Exit Function
ErrHandler:
5         Throw "#CDSRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


