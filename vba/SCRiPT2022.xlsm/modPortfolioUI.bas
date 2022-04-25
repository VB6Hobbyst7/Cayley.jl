Attribute VB_Name = "modPortfolioUI"
'---------------------------------------------------------------------------------------
' Module    : modPortfolioUI
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Code that implements the user interface o the Portfolio sheet, but see also
'             modDoubleclick, modUpdatePortfolioSheet and modSolve.
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 1

'---------------------------------------------------------------------------------------
' Procedure : CheckHeaders
' Author    : Philip Swannell
' Date      : 17-Nov-2015
' Purpose   : Checks that the headers on the Portfolio sheet are in synch with the headers
'             on the Hidden sheet. Throws an error if not since in that case it's probably
'             an error to copy and paste from the Hidden sheet to the Portfolio sheet.
'---------------------------------------------------------------------------------------
Sub CheckHeaders()
          Dim Headers1 As Range
          Dim Headers2 As Range
          Dim NumCols As Long

1         On Error GoTo ErrHandler
2         NumCols = RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Columns.Count

3         Set Headers1 = RangeFromSheet(shPortfolio, "PortfolioHeader").Resize(1, NumCols)
4         Set Headers2 = RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Cells(0, 1).Resize(1, NumCols)
5         If Not sArraysIdentical(Headers1, Headers2) Then
6             Throw "Assertion Failed: Headers on Portfolio sheet do not match Headers on template ranges on Hidden sheet"
7         End If

8         Exit Sub
ErrHandler:
9         Throw "#CheckHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CalculatePortfolioSheet
' Author    : Philip Swannell
' Date      : 13-Apr-2016
' Purpose   : Calculates the portfolio sheet without triggering its Worksheet_Calculate code
'---------------------------------------------------------------------------------------
Sub CalculatePortfolioSheet()
          Dim CopyOfErr As String
          Dim OldBCE As Boolean
1         On Error GoTo ErrHandler
2         OldBCE = gBlockCalculateEvent

3         gBlockCalculateEvent = True
4         shPortfolio.Calculate
5         gBlockCalculateEvent = OldBCE
6         Exit Sub
ErrHandler:
7         CopyOfErr = "#CalculatePortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
8         gBlockCalculateEvent = OldBCE
9         Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTradesRangeColumnWidths
' Author    : Philip
' Date      : 01-Apr-2016
' Purpose   : As per its name
'---------------------------------------------------------------------------------------
Sub SetTradesRangeColumnWidths()
          Dim FiltersAreBlank As Boolean
          Dim i As Long
          Dim MinimumWidths As Variant
          Dim NumTrades As Long
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradesRange As Range
          Const ExtraWidth = 0.5
          Const MaxWidth = 30

1         On Error GoTo ErrHandler

2         Set SPH = CreateSheetProtectionHandler(shPortfolio)
3         Set SUH = CreateScreenUpdateHandler()
4         Set TradesRange = getTradesRange(NumTrades)
5         MinimumWidths = RangeFromSheet(shHiddenSheet, "MinimumWidths").Value2

          'Setting the column widths is made more complex by the presence of merged cells between the filters and the _
           data, if they were not present we could simply call the AutoFitColumns method on the union of the filters range and the data range
6         With RangeFromSheet(shPortfolio, "TheFilters")
7             If Not IsEmpty(.Cells(1, 1)) Then
8                 FiltersAreBlank = False
9             Else
10                FiltersAreBlank = .Cells(1, 1).End(xlToRight).Column > .Column + .Columns.Count - 1
11            End If
12            If Not FiltersAreBlank Then
13                AutoFitColumns .Offset(0), ExtraWidth, MinimumWidths, MaxWidth, 1
14                For i = 1 To TradesRange.Columns.Count
15                    MinimumWidths(1, i) = MyMax(MinimumWidths(1, i), .Cells(1, i).ColumnWidth - ExtraWidth)
16                Next i
17            End If
18        End With

19        AutoFitColumns Application.Union(TradesRange.Rows(0).Resize(NumTrades + 1), RangeFromSheet(shPortfolio, "TotalPV")), ExtraWidth, MinimumWidths, MaxWidth, 12.43

20        Exit Sub
ErrHandler:
21        Throw "#SetTradesRangeColumnWidths (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateTradeTemplates
' Author    : Philip Swannell
' Date      : 11-Apr-2016
' Purpose   : Make the dates in the template trades suitable for new trades, currently
'             just doing the dates, but could do more.
'---------------------------------------------------------------------------------------
Sub UpdateTradeTemplates()
          Dim AnchorDate As Long
          Dim EndDate As Long
          Dim N As Name
          Dim StartDate As Long
          Dim SUH As SolumAddin.clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         OpenMarketWorkbook
3         AnchorDate = RangeFromMarketDataBook("Config", "AnchorDate").Value2
4         StartDate = AnchorDate + Choose((AnchorDate Mod 7) + 1, 3, 2, 2, 2, 2, 4, 4)
5         EndDate = Application.WorksheetFunction.EDate(StartDate, 60)

6         Set SUH = CreateSheetProtectionHandler(shHiddenSheet)

7         For Each N In shHiddenSheet.Names
8             If Right(N.Name, 8) = "Template" Then
9                 With N.RefersToRange.Cells(1, gCN_StartDate)
10                    If IsNumberOrDate(.Value) Then .Value = StartDate
11                End With
12                With N.RefersToRange.Cells(1, gCN_EndDate)
13                    If IsNumberOrDate(.Value) Then .Value = EndDate
14                End With
15            End If
16        Next N

          'On market Swap rate
          Dim ATMSwapRate
          Dim ExampleSwap
17        ExampleSwap = RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Value2
18        ExampleSwap(1, gCN_Counterparty) = gWHATIF
          'PGS 17 Nov 2020 TradeSolver not implemented for Julia
          ' ATMSwapRate = TradeSolver(ExampleSwap, "Rate 1")
19        ATMSwapRate = Empty
          '  If IsNumber(ATMSwapRate) Then
20        RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Cells(1, gCN_Rate1).Value = ATMSwapRate
          '  End If

          'On Market Fx Forward and ATM Fx Option
          Dim ATMForward
          Dim ExampleFxForward
21        ExampleFxForward = RangeFromSheet(shHiddenSheet, "FxForwardTemplate").Value2
22        ExampleFxForward(1, gCN_Counterparty) = gWHATIF
          '  ATMForward = TradeSolver(ExampleFxForward, "Notional 2")
23        ATMForward = Empty
          ' If IsNumber(ATMForward) Then
24        RangeFromSheet(shHiddenSheet, "FxForwardTemplate").Cells(1, gCN_Notional2).Value = ATMForward
25        RangeFromSheet(shHiddenSheet, "FxOptionTemplate").Cells(1, gCN_Notional2).Value = ATMForward
          ' End If

26        Exit Sub
ErrHandler:
27        Throw "#UpdateTradeTemplates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddTrades
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Adds trades of a given type to the Portfolio sheet
'---------------------------------------------------------------------------------------
Sub AddTrades(TradeType As String, NumTradesToAdd As Long)
          Dim CopyOfErr As String
          Dim DefaultNum As Long
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim NumExistingTrades As Long
          Dim OldBCE As Boolean
          Dim Res As Variant
          Dim SourceRange As Range
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim StateLimits As Boolean
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim TradesRange As Range
          Const MaxTrades = 10000

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         If NumTradesToAdd = 0 Then
TryAgain:
5             If TypeName(Selection) = "Range" Then
6                 DefaultNum = Selection.Areas(1).Rows.Count
7             Else
8                 DefaultNum = 1
9             End If

10            Res = InputBoxPlus("How many " + TradeType + "s do you want to add?" + IIf(StateLimits, vbLf + "Must be between 1 and " & Format(MaxTrades, "#,##0"), ""), MsgBoxTitle(), CStr(DefaultNum))
11            If VarType(Res) = vbBoolean Then Exit Sub
12            If Not IsNumeric(Res) Then GoTo TryAgain
13            NumTradesToAdd = CDbl(Res)

14            If NumTradesToAdd = 0 Then GoTo EarlyExit
15            If NumTradesToAdd < 1 Or NumTradesToAdd > MaxTrades Then
16                StateLimits = True
17                GoTo TryAgain
18            End If
19            If NumTradesToAdd <> CLng(NumTradesToAdd) Then GoTo TryAgain
20        End If

21        If NumTradesToAdd < 1 Then Throw "NumTradesToAdd must be positive"
22        If NumTradesToAdd > MaxTrades Then Throw "NumTradesToAdd must be less than or equal to " & Format(MaxTrades, "#,##0")

23        UpdateTradeTemplates

24        Set TradesRange = getTradesRange(NumExistingTrades)
25        Set SourceRange = RangeFromSheet(shHiddenSheet, TradeType + "Template")

26        If NumExistingTrades = 0 Then
27            Set TargetRange = TradesRange.Rows(1).Resize(NumTradesToAdd)
28        Else
29            Set TargetRange = TradesRange.Offset(TradesRange.Rows.Count).Resize(NumTradesToAdd)
30        End If

31        Set SPH = CreateSheetProtectionHandler(shPortfolio)
32        Set ExSH = CreateExcelStateHandler(PreserveViewport:=True)

33        If NumBlanksInRange(TargetRange) - TargetRange.Cells.Count > 1 Then    'Tolerance of one to cope with the possible presence of the "<Doubleclick to add trade>" cell
34            Application.GoTo TargetRange
35            Application.ScreenUpdating = True
36            If MsgBoxPlus("Overwrite these trades?", vbYesNoCancel + vbQuestion + vbDefaultButton2, MsgBoxTitle(), "Yes - Overwrite") <> vbYes Then GoTo EarlyExit
37        End If

38        Set SUH = CreateScreenUpdateHandler()

39        SourceRange.Copy
40        TargetRange.PasteSpecial xlPasteAll
41        CalculatePortfolioSheet

42        Application.CutCopyMode = False

43        FormatTradesRange
44        RepairTradeIDs
45        FilterTradesRange

46        CalculatePortfolioSheet
47        shxVADashboard.Calculate

48        Set ExSH = Nothing
49        If ActiveSheet Is shPortfolio Then
50            If Application.Intersect(TargetRange, ActiveWindow.VisibleRange) Is Nothing Then
51                Application.GoTo TargetRange.Offset(-9)
52                TargetRange.Select
53            Else
54                TargetRange.Select
55            End If
56        End If

EarlyExit:
57        gBlockChangeEvent = OldBCE

58        Exit Sub
ErrHandler:
59        CopyOfErr = "#AddTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
60        gBlockChangeEvent = OldBCE
61        Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : NumSelectedTrades
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Returns the number of trades that the user has selected
'---------------------------------------------------------------------------------------
Function NumSelectedTrades()
1         On Error GoTo ErrHandler
          Dim NumTrades
          Dim TradesRange As Range
2         If TypeName(Selection) <> "Range" Then
3             NumSelectedTrades = 0
4             Exit Function
5         ElseIf Not Selection.Parent Is shPortfolio Then
6             NumSelectedTrades = 0
7             Exit Function
8         Else
9             Set TradesRange = getTradesRange(NumTrades)
10            If NumTrades = 0 Then
11                NumSelectedTrades = 0
12                Exit Function
13            ElseIf Application.Intersect(Selection.EntireRow, TradesRange) Is Nothing Then
14                NumSelectedTrades = 0
15                Exit Function
16            End If
17            NumSelectedTrades = Application.Intersect(Selection.EntireRow, TradesRange.Columns(1)).Cells.Count
18        End If

19        Exit Function
ErrHandler:
20        Throw "#NumSelectedTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetTradesSummary
' Author     : Philip Swannell
' Date       : 05-Dec-2017
' Purpose    : Returns a "summary of the trades on the Portfolio sheet.
'              example: "20 x Swaption, 10 x InterestRateSwap, 1 x FxForward"
' Parameters :
'  Trades:    Trades in the same format as they appear on the Portfolio sheet, OR in the Julia trades format
' -----------------------------------------------------------------------------------------------------------------------
Function GetTradesSummary(Optional Trades)
          Dim ColNo
          Dim i As Long
          Dim NumTrades As Long
          Dim Result As String
          Dim VFs As Variant

1         On Error GoTo ErrHandler
2         If IsMissing(Trades) Or IsEmpty(Trades) Then
3             Trades = getTradesRange(NumTrades).Value2
4         End If

5         ColNo = sMatch("ValuationFunction", sArrayTranspose(sSubArray(Trades, 1, 1, 1)))

6         If IsNumber(ColNo) Then
7             NumTrades = sNRows(Trades) - 1
8             VFs = sSubArray(Trades, 2, ColNo, , 1)
9         Else

10            NumTrades = sNRows(Trades)
11            If NumTrades = 1 Then If IsEmpty(Trades(1, 1)) Then NumTrades = 0
12            VFs = sSubArray(Trades, 1, gCN_TradeType, , 1)
13            If NumTrades = 1 Then Force2DArray VFs
14        End If

15        If NumTrades = 0 Then
16            GetTradesSummary = "No Trades"
17            Exit Function
18        End If

19        VFs = sSortedArray(sCountRepeats(sSortedArray(VFs), "HC"), 1, , , False)
20        For i = 1 To sNRows(VFs)
21            Result = Result + IIf(i > 1, ", ", "") & Format(VFs(i, 1), "###,##0") & " x " & VFs(i, 2)
22        Next
23        GetTradesSummary = Result
24        Exit Function
ErrHandler:
25        Throw "#getTradesSummary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TradeFileInfo
' Author     : Philip Swannell
' Date       : 07-Feb-2019
' Purpose    : Gets either the TradeSummary or the CheckSum from a file, handles both types of trade file and caches results for speed
'              in the case of stf files, if the file (actually a workbook) does not have a Summary sheet then one is added and the file written back to disk
' Parameters :
'  FileName: file name with path
'  InfoType:  Either "TradesSummary" or "CheckSum"
' -----------------------------------------------------------------------------------------------------------------------
Function TradeFileInfo(FileName As String, InfoType As String)
          Static c As Collection
          Dim CheckSum As String
          Dim EXH As SolumAddin.clsExcelStateHandler
          Dim KeyCheckSum As String
          Dim KeyToUse As String
          Dim KeyTradesSummary As String
          Dim LMD As String
          Dim SheetName As String
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradesSummary As String
          Const AddressTradesSummary = "A1"
          Const AddressCheckSum = "A2"
          
1         On Error GoTo ErrHandler
          
2         LMD = CStr(CDbl(sFileLastModifiedDate(FileName)))
3         KeyTradesSummary = Replace(FileName + "!" + LMD, " ", "_") + "!" + "TradesSummary"
4         KeyCheckSum = Replace(FileName + "!" + LMD, " ", "_") + "!" + "CheckSum"
          
5         Select Case InfoType
              Case "TradesSummary"
6                 SheetName = "Summary"
7                 KeyToUse = KeyTradesSummary
8             Case "CheckSum"
9                 SheetName = "Summary"
10                KeyToUse = KeyCheckSum
11            Case Else
12                Throw "InfoType must be either 'TradesSummary' or 'CheckSum'"
13        End Select

14        If c Is Nothing Then Set c = New Collection

15        If IsInCollection(c, KeyToUse) Then
16            TradeFileInfo = c.item(KeyToUse)
17            Exit Function
18        End If
            
19        If LCase(Right(FileName, 4)) = ".stf" Then
20            TradesSummary = sCellContentsFromFileOnDisk(FileName, SheetName, AddressTradesSummary)
21            CheckSum = sCellContentsFromFileOnDisk(FileName, SheetName, AddressCheckSum)
            
22            If sIsErrorString(TradesSummary) Or sIsErrorString(CheckSum) Or Len(TradesSummary) < 2 Or Len(CheckSum) < 5 Then

23                If IsInCollection(Application.Workbooks, sSplitPath(FileName)) Then
24                    Throw "Cannot get " + InfoType + " for file '" + FileName + "' because a file with the same name is open in Excel"
25                End If
                  Dim wb As Workbook
                  Dim ws As Worksheet
26                Set SUH = CreateScreenUpdateHandler()
27                Set EXH = CreateExcelStateHandler(, , False)
28                Application.DisplayAlerts = False
29                Set wb = Application.Workbooks.Open(FileName)
30                If IsInCollection(wb.Worksheets(1).Names, "TradesNoHeaders") Then
31                    TradesSummary = GetTradesSummary(wb.Worksheets(1).Range("TradesNoHeaders"))
32                    CheckSum = sArrayCheckSum(wb.Worksheets(1).Range("TradesNoHeaders"))
33                Else
34                    CheckSum = ""
35                End If
36                If Not IsInCollection(wb.Worksheets, SheetName) Then
37                    Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
38                    ws.Name = SheetName
39                Else
40                    Set ws = wb.Worksheets(SheetName)
41                End If

42                ws.Range(AddressTradesSummary).Value = TradesSummary
43                ws.Range(AddressCheckSum) = CheckSum

44                On Error Resume Next
45                wb.Save
46                wb.Close False
47                On Error GoTo ErrHandler
48            End If
49        Else
50            CheckSum = sFileCheckSum(FileName)
              'it's presumably a file in the Julia format
              Dim TradesJuliaFormat
51            TradesJuliaFormat = sFileShow(FileName, ",", True, True, True, , "yyyy-mm-dd")

52            If Not sIsErrorString(TradesJuliaFormat) Then
53                TradesSummary = GetTradesSummary(TradesJuliaFormat)
54            Else
55                TradesSummary = "#Error getting trades - " + TradesJuliaFormat + "!"
56            End If
57        End If

58        c.Add CStr(TradesSummary), KeyTradesSummary
59        c.Add CStr(CheckSum), KeyCheckSum
          
60        Select Case InfoType
              Case "TradesSummary"
61                TradeFileInfo = CStr(TradesSummary)
62            Case "CheckSum"
63                TradeFileInfo = CStr(CheckSum)
64        End Select

65        Exit Function
ErrHandler:
66        TradeFileInfo = "#TradeFileInfo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' ----------------------------------------------------------------
' Procedure Name: AbbreviateTradeSummary
' Purpose: Translate a long-form trade summary e.g. "10 x FxOption, 10 x InterestRateSwap, 1 x CrossCurrencySwap"
'                              to a short-form e.g. "10FxO, 10IRS, 1CCS"
' Parameter Summary (String):
' Author: Philip Swannell
' Date: 05-Dec-2017
' ----------------------------------------------------------------
Function AbbreviateTradeSummary(Summary As String)
          Dim Res As String
1         On Error GoTo ErrHandler
2         Res = Summary
3         Res = Replace(Res, "CrossCurrencySwap", "CCS")
4         Res = Replace(Res, "InterestRateSwap", "IRS")
5         Res = Replace(Res, "InflationYoYSwap", "InfYoY")
6         Res = Replace(Res, "InflationZCSwap", "InfZC")
7         Res = Replace(Res, "FixedCashflows", "FC")
8         Res = Replace(Res, "FxOptionStrip", "FxOs")
9         Res = Replace(Res, "FxForwardStrip", "FxFs")
10        Res = Replace(Res, "FxOption", "FxO")
11        Res = Replace(Res, "CapFloor", "CF")
12        Res = Replace(Res, "Swaption", "Swptn")
13        Res = Replace(Res, "FxForward", "FxF")
14        Res = Replace(Res, " x ", "")
15        Res = Replace(Res, ", ", ",")
16        AbbreviateTradeSummary = Res
17        Exit Function
ErrHandler:
18        Throw "#AbbreviateTradeSummary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : getTradesRange
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Returns the range containing the trades, excludes header rows, includes
'             the three columns of formulas on the right, unless withCalculatedValues is passed as false
'---------------------------------------------------------------------------------------
Function getTradesRange(ByRef NumTrades, Optional withCalculatedValues As Boolean = True) As Range
          Static NumColsWith As Long
          Static NumColsWithout As Long
          Dim NC As Long

          Static HaveCheckedHeaders As Boolean
          Dim TheRange As Range
1         On Error GoTo ErrHandler

2         If NumColsWith = 0 Then
3             NumColsWith = RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Columns.Count
4         End If
5         If NumColsWithout = 0 Then
6             NumColsWithout = gCN_Counterparty
7         End If
8         NC = IIf(withCalculatedValues, NumColsWith, NumColsWithout)

9         If Not HaveCheckedHeaders Then
10            CheckHeaders
11            HaveCheckedHeaders = True
12        End If

          '13        Set SPH = CreateSheetProtectionHandler(shPortfolio)    'Since CurrentRegion will not run on protected sheet
13        Set TheRange = sExpandDown(RangeFromSheet(shPortfolio, "PortfolioHeader").Resize(, NC))
14        With TheRange
15            If .Cells(.Rows.Count, 2) = gDoubleClickPrompt Then
16                Set TheRange = .Resize(.Rows.Count - 1)
17            End If
18        End With

19        If TheRange.Row + TheRange.Rows.Count - 1 < RangeFromSheet(shPortfolio, "PortfolioHeader").Row + 1 Then
20            NumTrades = 0
21            Set TheRange = RangeFromSheet(shPortfolio, "PortfolioHeader").Cells(2, 1).Resize(, NC)
22        Else
23            Set TheRange = Range(RangeFromSheet(shPortfolio, "PortfolioHeader").Cells(2, 1), TheRange.Cells(TheRange.Rows.Count, 1)).Resize(, NC)
24            NumTrades = TheRange.Rows.Count
25        End If
26        Set getTradesRange = TheRange
27        With RangeFromSheet(shPortfolio, "NbTrades")
28            If .Value <> NumTrades Then
29                .Value = NumTrades
30            End If
31        End With

32        Exit Function
ErrHandler:
33        Throw "#getTradesRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function UnhiddenRowsInRange(TheRange As Range) As Range
          Dim a As Range
          Dim b As Boolean
          Dim CountRepeatsRet As Variant
          Dim i As Long
          Dim Indicators As Variant
          Dim Res As Range
          Dim ThisChunk As Range

1         On Error GoTo ErrHandler
2         For Each a In TheRange.Areas
3             Indicators = sReshape(True, a.Rows.Count, 1)
4             For i = 1 To a.Rows.Count
5                 Indicators(i, 1) = Not (a.Rows(i).Hidden)
6             Next i
7             CountRepeatsRet = sCountRepeats(Indicators, "CFH")
8             For i = 1 To sNRows(CountRepeatsRet)
9                 If CountRepeatsRet(i, 1) Then
10                    Set ThisChunk = a.Rows(CountRepeatsRet(i, 2)).Resize(CountRepeatsRet(i, 3))
11                    If Not b Then
12                        Set Res = ThisChunk
13                        b = True
14                    Else
15                        Set Res = Application.Union(Res, ThisChunk)
16                    End If
17                End If
18            Next i
19        Next a
20        Set UnhiddenRowsInRange = Res
21        Exit Function
ErrHandler:
22        Throw "#UnhiddenRowsInRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CloseTrades
' Author    : Philip Swannell
' Date      : 4 Feb 2019
' Purpose   : Deletes all trades from the Portfolio sheet.
'---------------------------------------------------------------------------------------
Sub CloseTrades()
          Dim BackupFileName As String
          Dim CopyOfErr As Boolean
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim NumTrades
          Dim OldBCE As Boolean
          Dim OldFileName As String
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradesRange As Range

1         On Error GoTo ErrHandler

2         Set TradesRange = getTradesRange(NumTrades, False)
3         If NumTrades = 0 Then Exit Sub

4         OldFileName = RangeFromSheet(shPortfolio, "TradesFileName")
5         Set ExSH = CreateExcelStateHandler(, , False, , , True)

6         Set SUH = CreateScreenUpdateHandler()
7         Set SPH = CreateSheetProtectionHandler(shPortfolio)
8         BackUpTrades False, BackupFileName      ' in case the user wishes they didn't...

9         If OldFileName <> "" Then
10            If TradeFileInfo(OldFileName, "CheckSum") <> sArrayCheckSum(TradesRange) Then
                  Dim MsgBoxRes As VbMsgBoxResult
                  Dim Prompt As String
11                Prompt = "Save your changes to '" + OldFileName + "'?"
12                MsgBoxRes = MsgBoxPlus(Prompt, vbExclamation + vbYesNoCancel, MsgBoxTitle(), "Save", "Don't Save", "Cancel")
13                If MsgBoxRes = vbCancel Then
14                    Exit Sub
15                ElseIf MsgBoxRes = vbYes Then
16                    ThrowIfError sFileCopy(BackupFileName, OldFileName)
17                End If
18            End If
19        End If

20        OldBCE = gBlockChangeEvent
21        gBlockChangeEvent = True

          Dim RangeToBackUp As Range
22        Set RangeToBackUp = getTradesRange(NumTrades)
23        Set RangeToBackUp = RangeToBackUp.Resize(RangeToBackUp.Rows.Count + 1)
24        BackUpRange RangeToBackUp, shUndo, Selection(), False

25        TradesRange.EntireRow.Delete
26        FilterTradesRange        'updates message "All xx trades shown."
27        CalculatePortfolioSheet
28        shxVADashboard.Calculate
29        SetTradesRangeColumnWidths
30        FilterTradesRange
31        RangeFromSheet(shPortfolio, "TradesFileName").ClearContents

32        gBlockChangeEvent = OldBCE

33        Set ExSH = Nothing        'restores viewport

34        Dim Res: Res = shPortfolio.UsedRange.Rows.Count    'resets UsedRange
35        Set SPH = Nothing
36        If Not RangeToBackUp Is Nothing Then
37            Application.OnUndo IIf(NumTrades = 1, "Restore deleted trade", "Restore " + Format(NumTrades, "###,##0") + " deleted trades"), "RestoreRange"
38        End If

39        Exit Sub
ErrHandler:
40        CopyOfErr = "#CloseTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
41        gBlockChangeEvent = OldBCE
42        Throw CopyOfErr
End Sub

'' -----------------------------------------------------------------------------------------------------------------------
'' Procedure  : CheckSumFromTradesFile
'' Author     : Philip Swannell
'' Date       : 06-Feb-2019
'' Purpose    : Returns the check sum for the trades in a file
'' Parameters :
''  FileName              : Full file name, with path.
''  FromScratchIfNecessary: If True then if the file does not contain a checksum then it is opened and the checksum is
''            recalculated from the trade data and the file is saved beck with the checksum embedded
''            (at cell A2 of sheet Summary).
'' -----------------------------------------------------------------------------------------------------------------------
'Function CheckSumFromTradesFile(FileName As String, FromScratchIfNecessary As Boolean) As String
'          Dim XSH As SolumAddin.clsExcelStateHandler, SPH As clsScreenUpdateHandler
'          Dim wb As Workbook
'          Dim ws As Worksheet
'          Dim Data As Variant
'          Const CellAddress = "A2" 'address of checksum on Summary sheet
'
'1         On Error GoTo ErrHandler
'2         If sFileExists(FileName) Then
'3             CheckSumFromTradesFile = sCellContentsFromFileOnDisk(FileName, "Summary", CellAddress)
'
'4             If Len(CheckSumFromTradesFile) > 10 Then Exit Function
'
'5             If FromScratchIfNecessary Then
'
'6                 Set XSH = CreateExcelStateHandler(, , False)
'7                 On Error Resume Next
'8                 Set wb = Application.Workbooks.Open(FileName)
'9                 Data = wb.Worksheets(1).Range("TradesNoHeaders").Value
'10                CheckSumFromTradesFile = sArrayCheckSum(Data)
'11                Set ws = wb.Worksheets("Summary")
'12                Set SPH = CreateSheetProtectionHandler(ws)
'13                ws.Range(CellAddress).Value = CheckSumFromTradesFile
'14                wb.Save 'attempt to save back the file, but with checksum written in
'15                wb.Close False
'16            Else
'17                CheckSumFromTradesFile = "#Checksum not found!"
'18            End If
'19        End If
'
'20        Exit Function
'ErrHandler:
'21        throw "#CheckSumFromTradesFile (line " & CStr(Erl) + "): " & Err.Description & "!"
'End Function

'---------------------------------------------------------------------------------------
' Procedure : DeleteSelectedTrades
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Deletes trades from the Portfolio shete that intersect the selected cells
'---------------------------------------------------------------------------------------
Sub DeleteSelectedTrades()
          Dim CopyOfErr As Boolean
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim Message As String
          Dim N As Long
          Dim NumTrades
          Dim OldBCE As Boolean
          Dim RangeToHighlight As Range
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradesRange As Range

1         On Error GoTo ErrHandler

2         Set TradesRange = getTradesRange(NumTrades)
3         If NumTrades = 0 Then Exit Sub

4         Set ExSH = CreateExcelStateHandler(, , , , , True)

5         Set RangeToHighlight = Application.Intersect(Selection.EntireRow, TradesRange)

6         If Not RangeToHighlight Is Nothing Then
7             Set RangeToHighlight = UnhiddenRowsInRange(RangeToHighlight)
8             If Not RangeToHighlight Is Nothing Then
9                 RangeToHighlight.Select
10                N = NumSelectedTrades()
11                Message = "Delete " + IIf(N > 1, "these " + CStr(N) + " trades?", "this trade?")
12                Message = Message + vbLf + vbLf + "Restore via Ctrl + Z or else via Menu > Open Trades > Restore trades from backups..."

13                If MsgBoxPlus(Message, vbYesNoCancel + vbDefaultButton2 + vbQuestion, MsgBoxTitle(), "Yes - Delete") = vbYes Then
14                    Set SUH = CreateScreenUpdateHandler()
15                    Set SPH = CreateSheetProtectionHandler(shPortfolio)
16                    BackUpTrades        ' in case the user wishes they didn't...
17                    OldBCE = gBlockChangeEvent
18                    gBlockChangeEvent = True

                      Dim RangeToBackUp As Range
19                    Set RangeToBackUp = getTradesRange(NumTrades)
20                    Set RangeToBackUp = RangeToBackUp.Resize(RangeToBackUp.Rows.Count + 1)
21                    BackUpRange RangeToBackUp, shUndo, Selection(), False

22                    RangeToHighlight.EntireRow.Delete
23                    FilterTradesRange        'updates message "All xx trades shown."
24                    CalculatePortfolioSheet
25                    shxVADashboard.Calculate
26                    SetTradesRangeColumnWidths
27                    FilterTradesRange
28                    Set TradesRange = getTradesRange(NumTrades)
29                    If NumTrades = 0 Then
30                        RangeFromSheet(shPortfolio, "TradesFileName").ClearContents
31                    Else
32                        AddGreyBorders TradesRange        'Deleting the bottom trade or trades leaves the bottom border missing so fix up.
33                    End If
34                End If
35            End If
36        End If

37        gBlockChangeEvent = OldBCE

38        Set ExSH = Nothing        'restores viewport

39        Dim Res: Res = shPortfolio.UsedRange.Rows.Count    'resets UsedRange
40        Set SPH = Nothing
41        If Not RangeToBackUp Is Nothing Then
42            Application.OnUndo IIf(N = 1, "Restore deleted trade", "Restore " + Format(N, "###,##0") + " deleted trades"), "RestoreRange"
43        End If

44        Exit Sub
ErrHandler:
45        CopyOfErr = "#DeleteSelectedTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
46        gBlockChangeEvent = OldBCE
47        Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TradeIDsNeedRepairing
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : Returns True if duplicates exit, empties exist etc.
'---------------------------------------------------------------------------------------
Function TradeIDsNeedRepairing(Optional OfferToFix As Boolean) As Boolean
          Dim ExistingTradeIDs
          Dim N As Long
          Dim NumTrades As Long
          Dim Prompt As String
          Dim TradeIDRange As Range

1         On Error GoTo ErrHandler

2         Set TradeIDRange = getTradesRange(NumTrades).Columns(1)
3         If NumTrades = 0 Then
4             TradeIDsNeedRepairing = False
5             Exit Function
6         End If

7         N = TradeIDRange.Rows.Count
8         ExistingTradeIDs = TradeIDRange.Value2
9         If Not sColumnAnd(sArrayIsNonTrivialText(ExistingTradeIDs))(1, 1) Then
10            TradeIDsNeedRepairing = True
11        ElseIf sNRows(sRemoveDuplicates(ExistingTradeIDs, True)) < sNRows(ExistingTradeIDs) Then
12            TradeIDsNeedRepairing = True
13        ElseIf IsNumber(sMatch("NEW", ExistingTradeIDs)) Then
14            TradeIDsNeedRepairing = True
15        ElseIf IsNumber(sMatch(Empty, ExistingTradeIDs)) Then
16            TradeIDsNeedRepairing = True
17        Else
18            TradeIDsNeedRepairing = False
19        End If

20        If TradeIDsNeedRepairing Then
21            If OfferToFix Then
22                Prompt = "Invalid or duplicated TradeIDs exist. Repair them now?"
23                If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, MsgBoxTitle(), "Repair") = vbOK Then
24                    RepairTradeIDs
25                    TradeIDsNeedRepairing = False
26                End If
27            End If
28        End If

29        Exit Function
ErrHandler:
30        Throw "#TradeIDsNeedRepairing (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : RepairTradeIDs
' Author    : Philip Swannell
' Date      : 18-Nov-2015
' Purpose   : Fixes up the TradeIDs of trades held on the Portfolio sheet.
'          * TradeIDs that are blank or read "NEW" are allocated new trade IDs
'          * If repeated trade IDs exist the second and subsequent ones will be fixed up
'---------------------------------------------------------------------------------------
Sub RepairTradeIDs()
          Dim CopyOfErr As String
          Dim ExistingTradeIDs
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim N As Long
          Dim NeedToFix As Boolean
          Dim NewTradeIDs
          Dim OldBCE As Boolean
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradeIDRange As Range

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True
4         Set SPH = CreateSheetProtectionHandler(shPortfolio)
5         Set SUH = CreateScreenUpdateHandler()

6         Set TradeIDRange = getTradesRange(N).Columns(1)
7         If N = 0 Then GoTo EarlyExit

8         ExistingTradeIDs = TradeIDRange.Value2
9         MatchIDs = sMatch(ExistingTradeIDs, ExistingTradeIDs)

10        Force2DArrayRMulti MatchIDs, ExistingTradeIDs, NewTradeIDs

11        NewTradeIDs = sReshape("", N + 1, 1)
12        For i = 1 To N + 1
13            NewTradeIDs(i, 1) = "T" + Format(i, "000000")
14        Next i

15        NewTradeIDs = sCompareTwoArrays(ExistingTradeIDs, NewTradeIDs, "In2AndNotIn1")
16        NewTradeIDs = sDrop(NewTradeIDs, 1)
17        NewTradeIDs = sArrayRange(NewTradeIDs, sReshape(0, sNRows(NewTradeIDs), 1))
18        For i = 1 To sNRows(NewTradeIDs)
19            NewTradeIDs(i, 2) = CLng(Replace(NewTradeIDs(i, 1), "T", ""))
20        Next i
21        NewTradeIDs = sSortedArray(NewTradeIDs, 2, , , True)
22        j = 1
          'NB change in logic below is likely to require a change in method TradeIDsNeedRepairing
23        For i = 1 To N
24            With TradeIDRange.Cells(i, 1)
25                If VarType(.Value) <> vbString Then
26                    NeedToFix = True
27                ElseIf .Value = "NEW" Or IsEmpty(.Value) Or MatchIDs(i, 1) <> i Then
28                    NeedToFix = True
29                Else
30                    NeedToFix = False
31                End If
32                If NeedToFix Then
33                    .Value = NewTradeIDs(j, 1)
34                    j = j + 1
35                End If
36            End With
37        Next i

          'check it worked!
38        ExistingTradeIDs = TradeIDRange.Value2
39        If sNRows(ExistingTradeIDs) <> sNRows(sRemoveDuplicates(ExistingTradeIDs)) Then
40            Throw "Assertion Failed: TradeIDs are not unique after running method RepairTradeIDs"
41        End If

EarlyExit:
42        gBlockChangeEvent = OldBCE

43        Exit Sub
ErrHandler:
44        gBlockChangeEvent = OldBCE
45        CopyOfErr = "#RepairTradeIDs (line " & CStr(Erl) + "): " & Err.Description & "!"
46        Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OpenTradesFile
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : Open a file containing trades and either replace the trades on the Portfolio
'             sheet or append them. Supports both stf file format and no longer supports the Airbus "Extract from Calypso" format
'             Return from function is Boolean to indicate if opening of the file was successful
'---------------------------------------------------------------------------------------
Function OpenTradesFile(Optional FullFileName As String, Optional activatePortfolioSheet As Boolean, Optional Overwrite As Variant) As Boolean
          Dim AlreadyOpen As Boolean
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim FileFilter As String
          Dim FileName As String
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         Set ExSH = CreateExcelStateHandler(, , , , , True)

3         If FullFileName = "" Then
4             FileFilter = "All trade files (*.stf;*csv),*.stf;*.csv,Solum Trade File (*.stf),*.stf,Comma Separated Trade files (*.csv),*.csv,All files (*.*),*.*"
5             FullFileName = GetOpenFilenameWrap(gProjectName & "TradeFiles", FileFilter, , "Open Trade File")
6         End If
7         If FullFileName = "False" Then
8             OpenTradesFile = False
9             Exit Function
10        End If
11        If Not sFileExists(FullFileName) Then        'Can happen when MRU contains files that no longer exist
12            RemoveFileFromMRU gProjectName & "TradeFiles", FullFileName
13            Throw "Cannot access the file " + vbLf + "'" + FullFileName + "'" + vbLf + _
                  "perhaps the file no longer exists or has been moved or renamed.", True
14        End If

          Dim FileType As String
          Dim FirstLineOfFile
15        FirstLineOfFile = ThrowIfError(sFileShow(FullFileName, "", , , , , , , , , 1, 1, 1, 1))(1, 1)
16        If Left(FirstLineOfFile, 2) = "PK" Then
17            FileType = "STF"
18        ElseIf InStr(LCase(FirstLineOfFile), "valuationfunction") > 0 Then
19            FileType = "CSV"
20        Else
21            Throw "File '" + FullFileName + "' is not a trades file"
22        End If

23        If FileType = "STF" Then
24            FileName = sSplitPath(FullFileName)
25            If IsInCollection(Application.Workbooks, FileName) Then
26                Set wb = Application.Workbooks(FileName)
27                If LCase(wb.FullName) <> LCase(FullFileName) Then
28                    Throw "A file " + FileName + " is already open in Excel. You cannot open two files with the same name"
29                End If
30                AlreadyOpen = True
31            Else
32                Application.DisplayAlerts = False
33                Set SUH = CreateScreenUpdateHandler()
34                Set wb = Workbooks.Open(FullFileName)
35            End If
36            If IsInCollection(wb.Worksheets(1).Names, "TradesWithHeaders") Then
37                ProcessSTFFile wb, Overwrite
38                OpenTradesFile = True
39            Else
40                If Not AlreadyOpen Then
41                    wb.Close False
42                End If
43                RemoveFileFromMRU gProjectName & "TradeFiles", FullFileName
44                Throw FullFileName + " is not a valid Solum Trades File (stf)", True
45            End If

46        ElseIf FileType = "CSV" Then
              Dim JuliaTrades
              Dim Numeraire As String
              Dim PortfolioTrades
47            JuliaTrades = ThrowIfError(sFileShow(FullFileName, ",", True, True, True, False, "yyyy-mm-dd"))
48            LongsToDoubles JuliaTrades
49            DateStringsToDoubles JuliaTrades
50            Numeraire = NumeraireFromJuliaTrades(JuliaTrades)
51            PortfolioTrades = JuliaTradesToPortfolioTrades(JuliaTrades, "EUR")
52            PasteTradesToPortfolioSheet PortfolioTrades, FullFileName, Overwrite
53            OpenTradesFile = True
54        End If

55        If activatePortfolioSheet Then
56            Set ExSH = Nothing
57            shPortfolio.Activate
58        End If

59        Exit Function
ErrHandler:
60        Throw "#OpenTradesFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : NumeraireFromJuliaTrades
' Author     : Philip Swannell
' Date       : 27-Nov-2020
' Purpose    : gets theNumeraire, stored (shameful bodge) in the header row.
' Parameters :
'  JuliaTrades:
' -----------------------------------------------------------------------------------------------------------------------
Private Function NumeraireFromJuliaTrades(JuliaTrades)
          Dim i As Long
          Dim j As Long
          Dim Numeraire As String
1         On Error GoTo ErrHandler
2         For j = 1 To sNCols(JuliaTrades)
3             If Left(JuliaTrades(1, j), 10) = "Numeraire=" Then
4                 Numeraire = Mid(JuliaTrades(1, j), 11)
5                 If Len(Numeraire) <> 3 Then Throw "Cannot find Numeraire in file"
6                 NumeraireFromJuliaTrades = Numeraire
7                 Exit Function
8             End If
9         Next
          'Oh dear not present, but that only matters if FxOption or FxOptionStrip trades exist.
          'TODO Change how we represent FxOption trades to store in the trade the Numeraire!!!!
10        Numeraire = "EUR" 'Does not matter if there are no Fx Options...
          Dim VFColNum As Long
11        For j = 1 To sNCols(JuliaTrades)
12            If JuliaTrades(1, j) = "ValuationFunction" Then
13                VFColNum = j
14                Exit For
15            End If
16        Next
17        For i = 2 To sNRows(JuliaTrades)
18            Select Case JuliaTrades(i, VFColNum)
                  Case "FxOption", "FxOptionStrip"
19                    Throw "Cannot find Numeraire in file"
20            End Select
21        Next i
22        NumeraireFromJuliaTrades = Numeraire

23        Exit Function
ErrHandler:
24        Throw "#NumeraireFromJuliaTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LongsToDoubles
' Author     : Philip Swannell
' Date       : 27-Nov-2020
' Purpose    : Necessary since validation in method JuliaTradesToPortfolio trades assumes all numbers are doubles?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub LongsToDoubles(ByRef MyArray)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         NR = sNRows(MyArray)
3         NC = sNCols(MyArray)
4         For i = 1 To NR
5             For j = 1 To NC
6                 If VarType(MyArray(i, j)) = vbLong Then MyArray(i, j) = CDbl(MyArray(i, j))
7             Next
8         Next

9         Exit Sub
ErrHandler:
10        Throw "#LongsToDoubles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DateStringsToDoubles
' Author     : Philip Swannell
' Date       : 07-Oct-2021
' Purpose    : Painful. Csv trade files have ended up containing dates in yyyy-mm-dd format but with leading and trailing
'              double quote characters. sFileShow now wraps sCSVRead and that correctly interprets such fields as strings
'              whereas earlier versions of sFileShow would convert such fields to dates (if argument ShowDatesAsDates is TRUE)
'              similar to sCSVReads "Q" option for the ConvertTypes argument, but not the intention of the coder (me) of
'              those earlier versions of sFileShow.
' Parameters :
'  MyArray:
' -----------------------------------------------------------------------------------------------------------------------
Private Sub DateStringsToDoubles(ByRef MyArray)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         NR = sNRows(MyArray)
3         NC = sNCols(MyArray)
4         For i = 1 To NR
5             For j = 1 To NC
6                 If VarType(MyArray(i, j)) = vbString Then
7                     If Len(MyArray(i, j)) = 10 Then
8                         If Mid$(MyArray(i, j), 5, 1) = "-" Then
9                             If Mid$(MyArray(i, j), 8, 1) = "-" Then
10                                MyArray(i, j) = CDbl(CDate(MyArray(i, j)))
11                            End If
12                        End If
13                    End If
14                End If
15            Next
16        Next

17        Exit Sub
ErrHandler:
18        Throw "#DateStringsToDoubles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ProcessSTFFile
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : Takes an STF file (a workbook) morphs the data in it for backward compatibility
'             then call PasteTradesToPortfolioSheet and closes the workbook
'---------------------------------------------------------------------------------------
Function ProcessSTFFile(wb As Workbook, Optional Overwrite As Variant)
          Dim ExistingHeaders As Range
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim NewHeaders
          Dim NumExistingTrades As Long
          Dim SourceRange As Range
          Dim SPH2 As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim ws As Worksheet
          Const SCRiPTWorkbookVersionOldestICanRead = 100
          Dim CopyOfErr As String
          Dim DataToPaste As Variant
          Dim FileName As String
          Dim i As Long
          Dim OldBCE As Boolean

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         getTradesRange NumExistingTrades

5         Set SUH = CreateScreenUpdateHandler()
6         Set ExSH = CreateExcelStateHandler(PreserveViewport:=True)
7         Set ws = wb.Worksheets(1)

          'Test for file being "too modern"
          Dim SCRiPTWorkBookVersionRequiredToRead As Long
8         If IsInCollection(ws.Names, "SCRiPTWorkBookVersionRequiredToRead") Then
9             SCRiPTWorkBookVersionRequiredToRead = Evaluate(ws.Names("SCRiPTWorkBookVersionRequiredToRead").Value)
10        Else
11            SCRiPTWorkBookVersionRequiredToRead = 142
12        End If
13        If SCRiPTWorkBookVersionRequiredToRead > SCRiPTWorkBookVersionNumber() Then
14            Throw "You need a more up-to-date version of the SCRiPT workbook to read that file. You have version " + CStr(SCRiPTWorkBookVersionNumber()) + " but you need version " + CStr(SCRiPTWorkBookVersionRequiredToRead) + " or later. "
15        End If

          'Test for file being "too ancient"
          Dim SavedWithSCRiPTWorkbookVersionNumber As Long
16        If IsInCollection(ws.Names, "SavedWithSCRiPTWorkbookVersionNumber") Then
17            SavedWithSCRiPTWorkbookVersionNumber = Evaluate(ws.Names("SavedWithSCRiPTWorkbookVersionNumber").Value)
18        Else
19            SavedWithSCRiPTWorkbookVersionNumber = 142
20        End If
21        If SCRiPTWorkbookVersionOldestICanRead > SavedWithSCRiPTWorkbookVersionNumber Then
22            Throw "That file was saved in a file format that's no longer supported by the SCRiPT workbook. It was saved with version " + CStr(SavedWithSCRiPTWorkbookVersionNumber) + " and this version of the SCRiPT can only read files saved with versions of at least " + CStr(SCRiPTWorkbookVersionOldestICanRead) + ". Sorry.", True
23        End If

24        If Not IsInCollection(ws.Names, "TradesWithHeaders") Then Throw "Cannot find Range name TradesWithHeaders in file opened"
25        If Not IsInCollection(ws.Names, "TradesNoHeaders") Then Throw "Cannot find Range name TradesNoHeaders in file opened"

26        Set ExistingHeaders = getTradesRange(NumExistingTrades).Offset(-2).Resize(2, gCN_Counterparty)
27        Set NewHeaders = RangeFromSheet(ws, "TradesWithHeaders").Resize(2)

          'Morphing...
          'As far as possible morph stf files that were saved with old versions of the SCRiPT workbook. So far (29 Feb 2016)
          'that merely means deleting any column headed "Trade Date", for which we have no use...
          'NB - ADDING NEW MORPHING CODE? _
           Then almost certainly the static SCRiPTWorkBookVersionRequiredToRead in method SaveTradesFile should be _
           updated to the current version number shown on the Audit sheet.
          Dim ColNo

28        ColNo = sMatch("Trade Date", sArrayTranspose(NewHeaders.Rows(2).Value))
29        If IsNumber(ColNo) Then
30            NewHeaders.Cells(1, ColNo).EntireColumn.Delete
31        End If
          'And morph "FIXED" > "Fixed", "FLOAT" > "Floating" and more...
32        Set SPH2 = CreateSheetProtectionHandler(ws)
          Dim c As Range
33        For Each c In RangeFromSheet(ws, "TradesNoHeaders").Cells
34            If Not c.HasFormula Then
35                Select Case CStr(c.Value)
                      Case "FIXED"
36                        c.Value = "Fixed"
37                    Case "FLOAT"
38                        c.Value = "Floating"
39                    Case "XCcySwap"
40                        c.Value = "CrossCurrencySwap"
41                    Case "FxFwd"
42                        c.Value = "FxForward"
43                    Case "Swap"
44                        c.Value = "InterestRateSwap"
45                End Select
46            End If
47        Next c
48        For i = 1 To 2
49            For Each c In RangeFromSheet(ws, "TradesNoHeaders").Columns(Choose(i, gCN_Freq1, gCN_Freq2)).Cells
50                Select Case c.Value
                      Case 1
51                        c.Value = "Annual"
52                    Case 2
53                        c.Value = "Semi annual"
54                    Case 4
55                        c.Value = "Quarterly"
56                    Case 12
57                        c.Value = "Monthly"
58                End Select
59            Next c
60        Next i

          '18 Nov 2020. Morph 'Is Fixed 1?' --> 'Leg Type 1'  and 'Is Fixed 2?' --> 'Leg Type 2' i.e. support trades from SCRiPT.xlsm
61        For i = 1 To 2
62            ColNo = sMatch("Is Fixed? " & CStr(i), sArrayTranspose(NewHeaders.Rows(2).Value))
63            If IsNumber(ColNo) Then
64                NewHeaders.Cells(2, ColNo).Value = "Leg Type " & CStr(i)
65                For Each c In RangeFromSheet(ws, "TradesNoHeaders").Columns(ColNo).Cells
66                    Select Case c.Value
                          Case "Floating"
67                            c.Value = "IBOR"
68                    End Select
69                Next c
70            End If
71        Next i

72        Set SPH2 = Nothing        'important to do this before closing the file

          'End of morphing

          Dim ExistingHeadersR2
          Dim NewHeadersR2
73        ExistingHeadersR2 = ExistingHeaders.Rows(2).Value
74        NewHeadersR2 = NewHeaders.Rows(2).Value
75        If sNCols(NewHeaders) <= 19 Then    'PGS 23 Jun 17. Made a small change to the headers. Counterparty was in _
                                               the top of the two rows, now it's in the bottom.
76            If NewHeaders(1, 19) = "" Then
77                NewHeaders(1, 19) = "Counterparty"
78            End If
79        End If

          'Check that headers have not changed.
80        If Not sArraysIdentical(ExistingHeadersR2, NewHeadersR2) Then Throw "Mismatch in header rows between data in file and data on the Portfolio sheet"
81        Set SourceRange = RangeFromSheet(ws, "TradesNoHeaders")

82        DataToPaste = SourceRange.Value2
83        FileName = wb.FullName
84        wb.Close False
85        PasteTradesToPortfolioSheet DataToPaste, FileName, Overwrite

86        gBlockChangeEvent = OldBCE
87        Exit Function
ErrHandler:
88        CopyOfErr = "#ProcessSTFFile (line " & CStr(Erl) + "): " & Err.Description & "!"
89        gBlockChangeEvent = OldBCE
90        If Not wb Is Nothing Then wb.Close False
91        Throw CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveTradesFile
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : Save an stf file (actually an excel file) to disk
'---------------------------------------------------------------------------------------
Function SaveTradesFile(Optional FileName As String, Optional ShowFileNameOnPortfolio = True, Optional AddToMRU As Boolean = True, Optional ByVal Validate As Boolean = True, Optional EvenWhenZeroTrades As Boolean)
          Dim ErrorMessage As String
          Dim NumTrades As Long
          Dim Prompt As String
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim TradesRange As Range
          Dim TradesRangeWithHeaders As Range
          Dim wb As Workbook
          Dim XSH As SolumAddin.clsExcelStateHandler
          'Files saved out by this method "know" what is the minimum version needed to read them!
          Const SCRiPTWorkBookVersionRequiredToRead = 457
          Dim CopyOfErr As String
          Dim OldBCE As Boolean
          Dim Trades

1         On Error GoTo ErrHandler
2         OldBCE = gBlockCalculateEvent
3         gBlockChangeEvent = True
            
4         Trades = getTradesRange(NumTrades).Value2
5         If NumTrades = 0 Then
6             Validate = False
7             If Not EvenWhenZeroTrades Then Throw "There are no trades to save", True
8         End If

9         If Validate Then
10            ValidateTrades Trades, False, ErrorMessage
11            If Len(ErrorMessage) > 0 Then
12                ErrorMessage = ErrorMessage + vbLf + vbLf + "Do you want to save the trades anyway?"
                  'Use vbOKCancel rather than vbYesNo so that Escape Key is equivalent to No
13                If MsgBoxPlus(ErrorMessage, vbOKCancel + vbQuestion + vbDefaultButton2, MsgBoxTitle(), "Yes, Save", "No, do nothing", , , 600) <> vbOK Then GoTo EarlyExit
14            End If
15        End If

16        If FileName = "" Then
17            FileName = GetSaveAsFilenameWrap(gProjectName & "TradeFiles", , "Solum Trade File,*.stf", , "Save Solum Trades File")
18            If FileName = "False" Then GoTo EarlyExit
19        End If

20        If Right(LCase(FileName), 4) <> ".stf" Then
21            If LCase(Right(FileName, 4)) = "xlsx" And InStr(LCase(sSplitPath(FileName)), "portfolio") > 0 Then
22                Prompt = "Cannot save in the current file format. Save in Solum Trades File format (.stf) instead?"
23            Else
24                Prompt = "Current file format not recognised. Save in Solum Trades File format (.stf) instead?"
25            End If
26            If MsgBoxPlus(Prompt, vbOKCancel + vbDefaultButton2 + vbExclamation, "Confirm Save Trades File", "Yes", "No") <> vbOK Then
27                GoTo EarlyExit
28            End If
              Dim DotAt
29            DotAt = InStrRev(sSplitPath(FileName), ".")
30            If DotAt = 0 Then
31                FileName = FileName + ".stf"
32            Else
33                FileName = Left(FileName, Len(FileName) - Len(sSplitPath(FileName)) + DotAt) + "stf"
34            End If
35        End If

36        Set SUH = CreateScreenUpdateHandler()

37        Set TradesRange = getTradesRange(NumTrades)

38        Set TradesRange = TradesRange.Resize(, gCN_Counterparty)

39        With TradesRange
40            Set TradesRangeWithHeaders = .Offset(-2).Resize(.Rows.Count + 2)
41        End With

42        Set XSH = CreateExcelStateHandler(, , False)

43        Set wb = Application.Workbooks.Add
44        If wb.Worksheets.Count < 2 Then wb.Worksheets.Add
45        With wb.Worksheets(2)
46            .Name = "Summary"
47            .Cells(1, 1).Value = GetTradesSummary()
48            .Cells(2, 1) = sArrayCheckSum(TradesRange.Value2)
49        End With
50        wb.Worksheets(1).Name = "Trades" 'PGS 7-Dec-2017. Trade files older than this date will have the name Sheet1 rather than Trades, SCRiPTWorkbookVersionNumber = 521

51        Set TargetRange = wb.Worksheets(1).Cells(1, 1).Resize(TradesRangeWithHeaders.Rows.Count, TradesRangeWithHeaders.Columns.Count)

52        TradesRangeWithHeaders.Copy
53        TargetRange.PasteSpecial xlPasteValues
54        Application.CutCopyMode = False

55        TargetRange.Worksheet.Names.Add "TradesWithHeaders", TargetRange
56        TargetRange.Worksheet.Names.Add "TradesNoHeaders", TargetRange.Offset(2).Resize(TargetRange.Rows.Count - 2)
          'Save versioning information
57        TargetRange.Worksheet.Names.Add "SavedWithSCRiPTWorkbookVersionNumber", SCRiPTWorkBookVersionNumber()
58        TargetRange.Worksheet.Names.Add "SCRiPTWorkBookVersionRequiredToRead", SCRiPTWorkBookVersionRequiredToRead

59        Application.DisplayAlerts = False
60        wb.SaveAs FileName, xlOpenXMLWorkbook, , , , False, , , False

61        If AddToMRU Then
62            AddFileToMRU gProjectName & "TradeFiles", FileName
63        Else
64            RemoveFileFromMRU gProjectName & "TradeFiles", FileName
65        End If

66        wb.Close False

67        Set SPH = CreateSheetProtectionHandler(shPortfolio)

68        If ShowFileNameOnPortfolio Then
69            RangeFromSheet(shPortfolio, "TradesFileName").Value = "'" + FileName
70        End If
71        TemporaryMessage "Trades saved to " + FileName

EarlyExit:
72        gBlockChangeEvent = OldBCE

73        Exit Function
ErrHandler:

74        CopyOfErr = "#SaveTradesFile (line " & CStr(Erl) + "): " & Err.Description & "!"
75        gBlockChangeEvent = OldBCE
76        Throw CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : SCRiPTWorkBookVersionNumber
' Author    : Philip Swannell
' Date      : 14-Mar-2016
' Purpose   : Returns the version number of this workbook from the workbook's Audit sheet.
'---------------------------------------------------------------------------------------
Function SCRiPTWorkBookVersionNumber()
1         SCRiPTWorkBookVersionNumber = RangeFromSheet(shAudit, "Headers").Cells(2, 1).Value
End Function


