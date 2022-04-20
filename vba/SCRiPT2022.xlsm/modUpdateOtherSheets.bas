Attribute VB_Name = "modUpdateOtherSheets"
Option Explicit

Sub ReloadResults()
1         On Error GoTo ErrHandler
5         Set gResults = JuliaExcel.JuliaUnserialiseFile(LocalTemp() & "results.txt", False)

8         Exit Sub
ErrHandler:
9         Throw "#ReloadResults (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Module    : modUpdateSheets
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Code to get data back from json output of the Julia code and paste it to the
'             xVADashboard, TradeViewer and CounterpartyViewer sheets.
'---------------------------------------------------------------------------------------
Sub ClearTradeViewerSheet()
          Dim SPH As SolumAddin.clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shTradeViewer)

3         With shTradeViewer
              Dim FirstCellToClear As Range
              Dim LastCellToClear As Range
              Dim WriteCell As Range
4             Set WriteCell = .Range("N9")
5             Set FirstCellToClear = .Cells(1, WriteCell.Column)
6             With .UsedRange
7                 Set LastCellToClear = .Cells(.Rows.Count, .Columns.Count)
8             End With
9             If LastCellToClear.Column < FirstCellToClear.Column Then
10                Set LastCellToClear = FirstCellToClear.Offset(, FirstCellToClear.Column - LastCellToClear.Column)
11            End If

12            Range(FirstCellToClear, LastCellToClear).Clear
13            .Calculate
14        End With

15        Exit Sub
ErrHandler:
16        Throw "#ClearTradeViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub TestUpdateTradeViewerSheet()
1         On Error GoTo ErrHandler
2         If gResults Is Nothing Then ReloadResults
3         If IsEmpty(gTradesAsOfLastPFECalc) Then gTradesAsOfLastPFECalc = getTradesRange(0).Value2
          
4         UpdateTradeViewerSheet True
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#TestUpdateTradeViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateTradeViewerSheet
' Author    : Philip Swannell
' Date      : 19-Apr-2016
' Purpose   : Update the TradeViewer sheet, which is now entirely free of formulas
'---------------------------------------------------------------------------------------
Sub UpdateTradeViewerSheet(Optional CallingFromMain As Boolean)
          Dim CashflowData As Variant
          Dim ChartTitle As String
          Dim i As Long
          Dim MatchID
          Dim OldBCE As Boolean
          Dim PFEData As Variant
          Dim PFEDataWithHeaders
          Dim TradeData
          Dim TradeID As String
          Dim ValuationFunction
          Dim xAxisMax As Double
          Dim xAxisMin As Double
          Const FirstWriteCell = "N9"

          Dim CopyOfErr As String
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         If Not IsPFEDataAvailable("trade", Not CallingFromMain) Then
5             If CallingFromMain Then
6                 ClearTradeViewerSheet
7                 GoTo EarlyExit
8             End If
9         End If

10        Set SPH = CreateSheetProtectionHandler(shTradeViewer)
11        Set SUH = CreateScreenUpdateHandler()

12        If CallingFromMain Then    'we force the tradeid to be valid
              Dim NewTradeID As Variant
              Dim ValidTradeIDs
13            ValidTradeIDs = sArrayTranspose(gResults("TradeExposures").keys)

14            If IsNumber(sMatch(CStr(RangeFromSheet(shTradeViewer, "SelectedTrade")), ValidTradeIDs)) Then
15                NewTradeID = RangeFromSheet(shTradeViewer, "SelectedTrade")
16            Else
17                NewTradeID = ValidTradeIDs(1, 1)
18            End If
19            With RangeFromSheet(shTradeViewer, "SelectedTrade")
20                .Value2 = NewTradeID
21                .NumberFormat = "General"
22            End With
23        End If

          'Get the blocks of data to paste
24        TradeID = RangeFromSheet(shTradeViewer, "SelectedTrade").Value2
25        PFEData = GetResultsFromJulia(TradeID, "Time,Date,TheirEPE,TheirENE,TheirPFE,TheirEE,OurSP", False, False)

26        If sIsErrorString(PFEData) Then
27            ClearTradeViewerSheet
28            shTradeViewer.Range(FirstWriteCell).Value = PFEData
29            GoTo EarlyExit
30        End If

31        PFEDataWithHeaders = sArrayStack(sArrayRange("Time", "Date", "EPE (CVA)", "ENE (DVA)", "PFE (KVA)", "EE (FVA)", "Surv Prob"), PFEData)
32        CashflowData = Empty

33        MatchID = sMatch(TradeID, sSubArray(gTradesAsOfLastPFECalc, 1, 1, , 1))
34        If IsNumber(MatchID) Then
35            TradeData = sSubArray(gTradesAsOfLastPFECalc, MatchID, 1, 1)
36            ValuationFunction = TradeData(1, gCN_TradeType)
37            If ValuationFunction = "InterestRateSwap" Or ValuationFunction = "CrossCurrencySwap" Or ValuationFunction = "CapFloor" Then
38                If Not IsEmpty(CashflowData) Then
39                    CashflowData = sSuppressNAs(CashflowData)
40                End If
41            End If
42        Else
43            TradeData = Empty
44            ValuationFunction = ""
45        End If

          Dim Paths
          Dim PathsAvailable As Boolean
46        Paths = GetResultsFromJulia(TradeID, "Time,Date,TheirPaths", False, False)
47        If sIsErrorString(Paths) Then
48            If InStr(Paths, "Paths not available") > 0 Then
49                PathsAvailable = False
50            Else
51                Throw Paths
52            End If
53        Else
54            PathsAvailable = True
55        End If

56        ClearTradeViewerSheet

          'Paste in trade data
57        With shTradeViewer
              Dim TradeLabelsRange As Range
              Dim TradeValuesRange As Range
              Dim WriteCell As Range
58            Set WriteCell = .Range(FirstWriteCell)

              'TODO the code here is very similar to a block of code in ViewTradeCashflows, put into a sub-routine and call from both places.
59            If Not IsEmpty(TradeData) Then
60                Set TradeLabelsRange = WriteCell.Resize(sNCols(TradeData))
61                Set TradeValuesRange = WriteCell.Offset(, 1).Resize(sNCols(TradeData))
62                TradeLabelsRange.Value2 = sArrayTranspose(RangeFromSheet(shHiddenSheet, "SingleRowHeaders").Value2)
63                TradeValuesRange.Value2 = sArrayTranspose(TradeData)

                  Dim TempRange As Range
64                With TradeValuesRange
65                    Set TempRange = .Cells(-1, 1).Resize(1, .Rows.Count)

66                    TempRange.Clear
67                    TempRange.Value2 = TradeData
68                    FormatTradesRange , TempRange
                      'using .Copy .PasteSpecial is incredibly slow, so I roll my own...
69                    For i = 1 To .Rows.Count
70                        TradeValuesRange.Cells(i, 1).NumberFormat = TempRange.Cells(1, i).NumberFormat
71                        TradeValuesRange.Cells(i, 1).Interior.Color = TempRange.Cells(1, i).Interior.Color
72                        TradeValuesRange.Cells(i, 1).Font.Color = TempRange.Cells(1, i).Font.Color
73                    Next i
74                    With TradeLabelsRange
75                        .Interior.Color = RGB(0, 102, 204)
76                        .Font.Color = RGB(255, 255, 255)
77                    End With
78                    With Application.Union(TradeLabelsRange, TradeValuesRange)
79                        .HorizontalAlignment = xlHAlignLeft
80                        AutoFitColumns .Offset(0), 0.5, , 20
81                        AddGreyBorders .Offset(0)
82                    End With
83                    TempRange.Clear
84                End With
85                Set WriteCell = WriteCell.Offset(0, 3)
86            End If

              'Paste in cashflow data
87            If Not IsEmpty(CashflowData) Then
                  Dim CashflowsRange As Range
88                Set CashflowsRange = WriteCell.Resize(sNRows(CashflowData), sNCols(CashflowData))
89                With CashflowsRange
90                    .Value2 = CashflowData
91                    ApplySolumFormatting .Offset(0), "Cashflows and PVs from " + CStr(ConfigRange("OurName").Value2) + "'s perspective.", ValuationFunction = ""
92                    Set WriteCell = WriteCell.Offset(, .Columns.Count + 1)
93                End With
94            End If

              'Paste in PFE data
              Dim Label As String
              Dim PFERange As Range
              Dim PFERangeNoHeaders As Range
              Dim TheirName As String
95            Set PFERange = WriteCell.Resize(sNRows(PFEDataWithHeaders), sNCols(PFEDataWithHeaders))
96            With PFERange
97                Set PFERangeNoHeaders = .Offset(1).Resize(.Rows.Count - 1)
98                If IsEmpty(TradeData) Then
99                    TheirName = "Bank"
100               Else
101                   TheirName = TradeData(1, gCN_Counterparty)
102                   If TheirName = gWHATIF Then TheirName = "Bank"
103               End If
104               Label = "Exposures from " + TheirName + "'s perspective."
105               .Value2 = PFEDataWithHeaders
106               ApplySolumFormatting .Offset(0), Label, False
107               SetCellComment .Cells(1, 3), StandardComment("EPE")
108               SetCellComment .Cells(1, 4), StandardComment("ENE")
109               SetCellComment .Cells(1, 5), StandardComment("PFE")
110               SetCellComment .Cells(1, 6), StandardComment("EE")
111               Set WriteCell = WriteCell.Offset(, .Columns.Count + 1)
112           End With

              'Paste in Paths
              Dim TargetRange As Range
113           If PathsAvailable Then
114               Set TargetRange = WriteCell.Offset(1).Resize(sNRows(Paths), sNCols(Paths))
115               With TargetRange
116                   .Value = Paths
117                   .Cells(0, 1) = "Time": .Cells(0, 2) = "Date"
118                   .Cells(0, 3).Resize(1, .Columns.Count - 2).Value = sArrayTranspose(sArrayConcatenate("Path ", sIntegers(.Columns.Count - 2)))
119                   Label = "Path values from " + TheirName + "'s perspective."
120                   ApplySolumFormatting .Offset(-1).Resize(.Rows.Count + 1), Label, False
121                   .Columns(1).NumberFormat = "0.000"
122                   .Columns(2).NumberFormat = NF_Date
123                   .Columns(3).Resize(, .Columns.Count - 2).NumberFormat = NF_Comma0dp
124                   AutoFitColumns .Offset(0), 0.5
125               End With
126           End If

              'Fix up the chart
127           With PFERangeNoHeaders
128               For i = 1 To 4
129                   SetChartData shTradeViewer.ChartObjects(1), i, .Cells(0, 2 + i), .Columns(1), .Columns(2 + i)
130               Next i
131           End With
132           ChartTitle = CStr(RangeFromSheet(shTradeViewer, "SelectedTrade")) + IIf(ValuationFunction <> "", " - " & ValuationFunction, "")
133           xAxisMin = 0
134           xAxisMax = EndOfNonZeroData(PFEData)
135           AmendChart shTradeViewer.ChartObjects(1), ChartTitle, CDbl(xAxisMin), xAxisMax, , , "0"
136       End With

EarlyExit:

137       gBlockChangeEvent = OldBCE
138       Exit Sub
ErrHandler:
139       CopyOfErr = "#UpdateTradeViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
140       gBlockChangeEvent = OldBCE
141       Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetCellComment
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : Adds a comment to a cell and makes it appear in Calibri 11. Comment must be
'             passed including line feed characters
'---------------------------------------------------------------------------------------
Function SetCellComment(c As Range, Comment As String)
1         On Error GoTo ErrHandler
2         c.ClearComments
3         c.AddComment
4         c.Comment.Visible = False
5         c.Comment.text text:=Comment
6         With c.Comment.Shape.TextFrame
7             .Characters.Font.Name = "Calibri"
8             .Characters.Font.Size = 11
9             .AutoSize = True
10        End With
11        Exit Function
ErrHandler:
12        Throw "#SetCellComment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : StandardComment
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : Avoid coding these strings in more than one place
'---------------------------------------------------------------------------------------
Private Function StandardComment(Topic As String) As String
1         Select Case Topic
              Case "EPE"
2                 StandardComment = "Expected Positive Exposure (EPE) from" + vbLf + "the Bank's perspective. EPE (plus the" + vbLf + "client's credit spreads) determines" + vbLf + "Credit Valuation Adjustment (CVA)"
3             Case "ENE"
4                 StandardComment = "Expected Negative Exposure (ENE)" + vbLf + "from the Bank's perspective. ENE (plus" + vbLf + "the Bank's credit spreads) determines" + vbLf + "Debt Valuation Adjustment (DVA)"
5             Case "PFE"
6                 StandardComment = "Potential Future Exposure (PFE) from" + vbLf + "the Bank's perspective. PFE is an" + vbLf + "important input to the calculation of" + vbLf + "Capital Valuation Adjustment (KVA)"
7             Case "EE"
8                 StandardComment = "Expected Exposure (EE) from the" + vbLf + "Bank's perspective. EE is used to" + vbLf + "calculate Funding Valuation" + vbLf + "Adjustment (FVA)"
9         End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : ApplySolumFormatting
' Author    : Philip Swannell
' Date      : 18-May-2016
' Purpose   : Appplies formatting to a range of cells where the top row is a header row in
'             "Solum Blue", remaining rows have light grey borders with number fomatting
'             guessed from the contents of the header cell. TODO: Move to SolumAddin.
'        If argument SuppressZeros is TRUE then number formats make cells containing zero
'             appear blank.
'---------------------------------------------------------------------------------------
Function ApplySolumFormatting(RangeWithHeaders As Range, Label As String, SuppressZeros As Boolean)
          Dim ColumnFormats As Variant
          Dim Header As String
          Dim i As Long
          Dim j As Long
          Dim NumberFormats(1 To 7) As String
          Dim rx(1 To 7) As New VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler

2         For i = UBound(rx) To LBound(rx)
3             rx(i).Global = False
4             rx(i).IgnoreCase = True
5         Next

6         rx(1).Pattern = "Date|AccFrom|AccTo"
7         rx(2).Pattern = "Notional|PV\(|Flow|PV \(|Amount|Annuity|CVA|DVA|FVA|KVA|FCA|FBA|Path"
8         rx(3).Pattern = "Forward\(.*/.*\)|Strike\(.*/.*\)"
9         rx(4).Pattern = "^Vol|Forward"
10        rx(5).Pattern = "Strike"
11        rx(6).Pattern = "^DCF|^DF"
12        rx(7).Pattern = "Time"

13        NumberFormats(1) = "dd-mmm-yyyy;;#"    'always suppress zeros in this case
14        NumberFormats(2) = IIf(SuppressZeros, NF_Comma0dp & ";#", NF_Comma0dp)
15        NumberFormats(3) = NF_Fx
16        NumberFormats(4) = IIf(SuppressZeros, "0.0000%;0.0000%;#", "0.0000%")
17        NumberFormats(5) = IIf(SuppressZeros, "0.00%;0.00%;#", "0.00%")
18        NumberFormats(6) = IIf(SuppressZeros, "0.00000;0.00000;#", "0.00000")
19        NumberFormats(7) = IIf(SuppressZeros, "0.000;0.000;#", "0.000")

20        With RangeWithHeaders
21            .ClearFormats
22            .HorizontalAlignment = xlHAlignCenter

23            If .Row > 1 Then
24                If Label <> "" Then
25                    With .Cells(0, 1)
26                        .Value = Label
27                        .Font.Color = Colour_GreyText
28                    End With
29                End If
30            End If

31            If .Column > 1 Then .Columns(0).ColumnWidth = 2
32            .Columns(.Columns.Count + 1).ColumnWidth = 2

33            With .Rows(1)
34                .Interior.Color = RGB(0, 102, 204)
35                .Font.Color = RGB(255, 255, 255)
36                .Font.Bold = True
37            End With

38            ColumnFormats = sReshape("General", .Columns.Count, 1)
39            For i = 1 To .Columns.Count
40                Header = CStr(.Cells(1, i).Value)
41                For j = 1 To 7
42                    If rx(j).Test(Header) Then
43                        ColumnFormats(i, 1) = NumberFormats(j)
44                        Exit For
45                    End If
46                Next j
47            Next i
              'Faster to apply number formatting in blocks, so use sCountRepeats trick
48            ColumnFormats = sCountRepeats(ColumnFormats, "CFH")
                  
49            For i = 1 To sNRows(ColumnFormats)
50                .Offset(, ColumnFormats(i, 2) - 1).Resize(, ColumnFormats(i, 3)).NumberFormat = ColumnFormats(i, 1)
51            Next i
                  
52            AddGreyBorders .Offset(0)
53            AutoFitColumns .Offset(0), 0.5
54            If SuppressZeros Then
55                For i = 1 To sNRows(ColumnFormats)
56                    If ColumnFormats(i, 1) = "General" Then
57                        .Offset(ColumnFormats(i, 2) - 1).Resize(, ColumnFormats(i, 3)).NumberFormat = "General;General;"    'This format suppresses display of zero, but interacts badly with auto-fitting of columns
58                    End If
59                Next i
60            End If

61        End With
62        Exit Function
ErrHandler:
63        Throw "#ApplySolumFormatting (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsPFEDataAvailable
' Author    : Philip Swannell
' Date      : 18-May-2016
' Purpose   : Determines whether the R environment holds PFE data, it may not if PFEs have
'             not been run or if at the last run there were no trades for the chosen banks and no WHATIF trades
'---------------------------------------------------------------------------------------
Function IsPFEDataAvailable(Level As String, ThrowIfNotAvailable As Boolean) As Boolean
          Dim DataExists As Boolean
          Dim ErrorString As String
          Dim PartitionByTrade As Variant
          Dim Res As Variant
          Const Error1 = "Trade level data is not available. Use Menu > Calculate xVAs with PFE and KVA (inc. by trade)"
          Const Error2 = "Trade level data is not available. To generate trade level data, first set PartitionByTrade to TRUE on the Config sheet and then generate data via Menu > Calculate xVAs with PFE and KVA (inc. by trade)"
          Const Error3 = "PFE data is not available. To generate the data use Menu > Calculate PVs, CVA and PFE"
          
1         On Error GoTo ErrHandler

          Dim PFEStatus

2         Select Case LCase(Level)
              Case "trade"
3                 On Error Resume Next
4                 PFEStatus = gResults("TradeResults")("PFEStatus")
5                 PartitionByTrade = gResults("Control")("PartitionByTrade")
6                 On Error GoTo ErrHandler
7                 If VarType(PartitionByTrade) <> vbBoolean Then PartitionByTrade = sEquals(ConfigRange("PartitionByTrade").Value, True)

8                 If PartitionByTrade Then
9                     ErrorString = Error1
10                Else
11                    ErrorString = Error2
12                End If
13            Case "netset"
14                On Error Resume Next
15                PFEStatus = gResults("PartyResults")("PFEStatus")
16                On Error GoTo ErrHandler

17                ErrorString = Error3
18            Case Else
19                Throw "Level must be 'Trade' or 'NetSet'"
20        End Select

21        DataExists = Not gResults Is Nothing
22        If DataExists Then
23            Res = sAny(sArrayEquals(PFEStatus, "OK"))
24            If VarType(Res) = vbBoolean Then
25                DataExists = Res
26            Else
27                DataExists = False
28            End If
29        End If

30        IsPFEDataAvailable = DataExists

31        If ThrowIfNotAvailable Then
32            If Not DataExists Then
33                Throw ErrorString, True
34            End If
35        End If

36        Exit Function
ErrHandler:
37        Throw "#IsPFEDataAvailable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub ClearCounterpartyViewerSheet()
          Dim SPH As SolumAddin.clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shCounterpartyViewer)
3         With shCounterpartyViewer
              'Clear out the sheet
              Dim b As Button
              Dim FirstCellToClear As Range
              Dim LastCellToClear As Range
              Dim WriteCell As Range
4             Set WriteCell = .Range("M10")
5             Set FirstCellToClear = .Cells(1, WriteCell.Column)
6             With .UsedRange
7                 Set LastCellToClear = .Cells(.Rows.Count, .Columns.Count)
8             End With
9             If LastCellToClear.Column < FirstCellToClear.Column Then
10                Set LastCellToClear = FirstCellToClear.Offset(, FirstCellToClear.Column - LastCellToClear.Column)
11            End If
12            Range(FirstCellToClear, LastCellToClear).Clear
13            For Each b In .Buttons
14                b.Delete
15            Next
16            .Calculate
17        End With
18        Exit Sub
ErrHandler:
19        Throw "#ClearCounterpartyViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TestUpdateCounterpartyViewerSheet
' Author    : Philip Swannell
' Date      : 03-Jan-2017
' Purpose   : Trivial test harness
'---------------------------------------------------------------------------------------
Sub TestUpdateCounterpartyViewerSheet()
1         On Error GoTo ErrHandler
2         If gResults Is Nothing Then ReloadResults
3         UpdateCounterpartyViewerSheet False
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestUpdateCounterpartyViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateCounterpartyViewerSheet
' Author    : Philip Swannell
' Date      : 19-Apr-2016
' Purpose   : Re-draws the sheet CounterpartyViewer
'---------------------------------------------------------------------------------------
Sub UpdateCounterpartyViewerSheet(Optional CallingFromRunRCode As Boolean)
          Dim ChartTitle As String
          Dim CopyOfErr As String
          Dim Counterparty As String
          Dim DataToNotPaste As Variant
          Dim i As Long
          Dim IncludeHypotheticals As Boolean
          Dim Label As String
          Dim maxY
          Dim minY
          Dim Numeraire As String
          Dim NumTrades As Long
          Dim OldBCE As Boolean
          Dim PFEData As Variant
          Dim PFEDataWithHeaders As Variant
          Dim PFERangeNoHeaders As Range
          Dim Res
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim TradeData As Variant
          Dim TradeDataWithHeaders As Variant
          Dim xAxisMax As Double
          Dim xAxisMin As Double
          Dim yAxisMax As Double
          Dim yAxisMin As Double
          Const ControlString = "Time,Date,TheirEPE,TheirENE,TheirPFE,TheirEE,OurSP"
          Const ControlStringWI = "Time,Date,TheirEPEWhatIf,TheirENEWhatIf,TheirPFEWhatIf,TheirEEWhatIf,OurSP"
          Const FirstWriteCell = "M10"

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         If Not IsPFEDataAvailable("netset", Not CallingFromRunRCode) Then
5             If CallingFromRunRCode Then
6                 ClearCounterpartyViewerSheet
7                 GoTo EarlyExit
8             End If
9         End If

10        If CallingFromRunRCode Then    'We force the Counterparty to be valid
              Dim NewCounterparty As String
              Dim ValidCounterparties
11            ValidCounterparties = CounterpartiesFromJulia()

12            If IsNumber(sMatch(CStr(RangeFromSheet(shCounterpartyViewer, "SelectedCpty")), ValidCounterparties)) Then
13                NewCounterparty = RangeFromSheet(shCounterpartyViewer, "SelectedCpty")
14            Else
15                NewCounterparty = ValidCounterparties(1, 1)
16            End If
17            RangeFromSheet(shCounterpartyViewer, "SelectedCpty").Resize(2).Value = sArrayStack(NewCounterparty, False)
18        End If

19        Counterparty = RangeFromSheet(shCounterpartyViewer, "SelectedCpty")
20        IncludeHypotheticals = RangeFromSheet(shCounterpartyViewer, "Inc_Hypotheticals?")
21        Numeraire = gResults("Model")("Numeraire")

22        PFEData = GetResultsFromJulia(Counterparty, IIf(IncludeHypotheticals, ControlStringWI, ControlString), False, True)
23        If sIsErrorString(PFEData) Then
24            ClearCounterpartyViewerSheet
25            Set SPH = CreateSheetProtectionHandler(shCounterpartyViewer)
26            shCounterpartyViewer.Range(FirstWriteCell).Value = PFEData
27            GoTo EarlyExit
28        End If
          
29        PFEDataWithHeaders = sArrayStack(sArrayRange("Time", "Date", "EPE (CVA)", "ENE (DVA)", "PFE (KVA)", "EE (FVA)", "Surv Prob"), PFEData)

30        TradeData = TradeDataForCounterpartyViewerSheet(Counterparty, IncludeHypotheticals, gWHATIF)
31        NumTrades = sNRows(TradeData) - 1
32        If NumTrades = 0 Then
33            TradeData = CreateMissing()
34        Else
35            TradeData = sSubArray(TradeData, 2, 1)        'Strip off the row names
36        End If

37        TradeDataWithHeaders = sArrayStack(sArrayRange("Trade ID", "Trade Type", "Counterparty", "Start Date", "End Date", "PV (in " & Numeraire & ")", "CVA", "DVA", "FCA", "FBA", "FVA"), TradeData)
38        DataToNotPaste = ThrowIfError(GetResultsFromJulia(Counterparty, IIf(IncludeHypotheticals, ControlString, ControlStringWI), False, True))
          
          Dim Paths
          Dim PathsAvailable As Boolean
39        Paths = GetResultsFromJulia(Counterparty, IIf(IncludeHypotheticals, "Time,Date,TheirPathsWhatIf", "Time,Date,TheirPaths"), False, True)

40        If sIsErrorString(Paths) Then
41            If InStr(Paths, "Paths not available") > 0 Then
42                PathsAvailable = False
43            Else
44                ClearCounterpartyViewerSheet
45                Set SPH = CreateSheetProtectionHandler(shCounterpartyViewer)
46                shCounterpartyViewer.Range(FirstWriteCell).Value = "Error getting Path data: " & Paths
47                GoTo EarlyExit
48            End If
49        Else
50            PathsAvailable = True
51        End If
          
52        Set SUH = CreateScreenUpdateHandler()
53        Set SPH = CreateSheetProtectionHandler(shCounterpartyViewer)
54        ClearCounterpartyViewerSheet

55        With shCounterpartyViewer
              Dim WriteCell As Range
56            Set WriteCell = .Range(FirstWriteCell)

              'Paste in the trade data
57            Set TargetRange = WriteCell.Resize(sNRows(TradeDataWithHeaders), sNCols(TradeDataWithHeaders))
58            With TargetRange
59                .Value2 = TradeDataWithHeaders
60                .Parent.Names.Add "TradeDataWithHeaders", .Offset(0)
61                If NumTrades > 0 Then
62                    With .Cells(-1, 6).Resize(1, 6)
63                        .Value = sColumnSum(sSubArray(TradeData, 1, 6, , 6))
64                        .NumberFormat = NF_Comma0dp
65                        .HorizontalAlignment = xlHAlignCenter
66                    End With
67                End If
68                ApplySolumFormatting .Offset(), "", False
                  'Have to do the autofit below because of the cells containing totals
69                AutoFitColumns .Offset(-2, 5).Resize(.Rows.Count + 2, 6), 0.5
70                AddSortButtons .Rows(0), 1
71                With .Cells(-1, 1)
72                    .Value = "Double-click to drill down"
73                    .Font.Color = Colour_GreyText
74                End With
75                SetCellComment .Cells(1, 6), "PV from Bank's" + vbLf + "perspective"
76                SetCellComment .Cells(1, 7), "For trade-by-trade CVA" + vbLf + "to be calculated, set" + vbLf + "PartitionByTrade to" + vbLf + "TRUE on the Config sheet."
77                SetCellComment .Cells(1, 8), "For trade-by-trade DVA" + vbLf + "to be calculated, set" + vbLf + "PartitionByTrade to" + vbLf + "TRUE on the Config sheet."
78                SetCellComment .Cells(1, 9), "For trade-by-trade FCA" + vbLf + "to be calculated, set" + vbLf + "PartitionByTrade to" + vbLf + "TRUE on the Config sheet."
79                SetCellComment .Cells(1, 10), "For trade-by-trade FBA" + vbLf + "to be calculated, set" + vbLf + "PartitionByTrade to" + vbLf + "TRUE on the Config sheet."

80                Set WriteCell = WriteCell.Offset(, .Columns.Count + 1)
81            End With

              'Paste in PFE data
82            Set TargetRange = WriteCell.Resize(sNRows(PFEDataWithHeaders), sNCols(PFEDataWithHeaders))
83            With TargetRange
84                Set PFERangeNoHeaders = .Offset(1).Resize(.Rows.Count - 1)
85                .Value = PFEDataWithHeaders
86                ApplySolumFormatting .Offset(0), "", False
87                SetCellComment .Cells(1, 3), StandardComment("EPE")
88                SetCellComment .Cells(1, 4), StandardComment("ENE")
89                SetCellComment .Cells(1, 5), StandardComment("PFE")
90                SetCellComment .Cells(1, 6), StandardComment("EE")
91                With .Rows(0)
92                    .Cells(1, 1).Value = IIf(IncludeHypotheticals, "Including Hypotheticals", "Base Case")
93                    .Font.Color = Colour_GreyText
94                End With
95                Set WriteCell = WriteCell.Offset(0, .Columns.Count + 1)
96            End With

              'Paste in Paths
97            If PathsAvailable Then
98                Set TargetRange = WriteCell.Offset(1).Resize(sNRows(Paths), sNCols(Paths))
99                With TargetRange
100                   .Value = Paths
101                   .Cells(0, 1) = "Time": .Cells(0, 2) = "Date"
102                   .Cells(0, 3).Resize(1, .Columns.Count - 2).Value = sArrayTranspose(sArrayConcatenate("Path ", sIntegers(.Columns.Count - 2)))
103                   Label = "Path values from " + Counterparty + "'s perspective."
104                   ApplySolumFormatting .Offset(-1).Resize(.Rows.Count + 1), Label, False
105               End With
106           End If
107       End With

          'Fix up the chart
108       With PFERangeNoHeaders
109           For i = 1 To 4
110               SetChartData shCounterpartyViewer.ChartObjects(1), i, .Cells(0, 2 + i), .Columns(1), .Columns(2 + i)
111           Next i
112       End With
113       maxY = MyMax(sMaxOfArray(sSubArray(DataToNotPaste, 1, 3)), sMaxOfArray(sSubArray(PFEData, 1, 3)))
114       minY = MyMin(sMinOfArray(sSubArray(DataToNotPaste, 1, 3)), sMinOfArray(sSubArray(PFEData, 1, 3)))
115       Res = ChartMaxMin(CDbl(minY), CDbl(maxY))
116       yAxisMin = Res(1, 1)
117       yAxisMax = Res(2, 1)
118       xAxisMax = MyMax(EndOfNonZeroData(PFEData), EndOfNonZeroData(DataToNotPaste))
119       xAxisMin = 0
120       ChartTitle = RangeFromSheet(shCounterpartyViewer, "SelectedCpty").Value & "'s exposure to " & _
              ConfigRange("OurName").Value & (IIf(IncludeHypotheticals, " with hypothetical trades", ""))
121       AmendChart shCounterpartyViewer.ChartObjects(1), ChartTitle, CDbl(xAxisMin), xAxisMax, yAxisMin, yAxisMax, "0"

EarlyExit:
122       gBlockChangeEvent = OldBCE
123       Exit Sub
ErrHandler:
124       CopyOfErr = "#UpdateCounterpartyViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
125       gBlockChangeEvent = OldBCE
126       Throw CopyOfErr
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetChartData
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : for an existing chart, amend the ranges from which data is taken for one
'             of the chart's series
'---------------------------------------------------------------------------------------
Function SetChartData(chOb As ChartObject, SeriesNumber As Long, SeriesName As Range, xValues As Range, Values As Range)
1         On Error GoTo ErrHandler
2         With chOb.Chart.FullSeriesCollection(SeriesNumber)
3             .xValues = "=" + xValues.Address(External:=True)
4             .Values = "=" + Values.Address(External:=True)
5             .Name = "=" + SeriesName.Address(External:=True)
6         End With
7         Exit Function
ErrHandler:
8         Throw "#SetChartData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : AmendChart
' Author    : Philip Swannell
' Date      : 29-Apr-2016
' Purpose   : Amend a chart to reflect passed in data
'---------------------------------------------------------------------------------------
Sub AmendChart(chOb As ChartObject, Title As String, xAxisMin As Double, xAxisMax As Double, Optional yAxisMin As Double, Optional yAxisMax As Double, Optional xAxisNumberFormat As String)
1         On Error GoTo ErrHandler
2         With chOb.Chart.Axes(xlCategory)
3             If .MinimumScale <> xAxisMin Then
4                 .MinimumScale = xAxisMin
5             End If
6             If .MaximumScale <> xAxisMax Then
7                 .MaximumScale = xAxisMax
8             End If
9             If xAxisNumberFormat <> "" Then
10                .TickLabels.NumberFormat = xAxisNumberFormat
11            End If
12        End With

13        If yAxisMin <> 0 Or yAxisMax <> 0 Then
14            With chOb.Chart.Axes(xlValue)
15                If .MinimumScale <> yAxisMin Then
16                    .MinimumScale = yAxisMin
17                End If
18                If .MaximumScale <> yAxisMax Then
19                    .MaximumScale = yAxisMax
20                End If
21            End With
22        Else
23            With chOb.Chart.Axes(xlValue)
24                .MinimumScaleIsAuto = True
25                .MaximumScaleIsAuto = True
26            End With
27        End If

28        chOb.Chart.ChartTitle.Caption = Title

29        Exit Sub
ErrHandler:
30        Throw "#AmendChart (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartMaxMin
' Author    : Philip Swannell
' Date      : 29-Apr-2016
' Purpose   : Determine appropriate values for Y-axis maximum and minimum on a chart
'             where the data to be plotted has max value DataMax and min value DataMin.
'             This is an attempt to replicate Excel's MaximumScalIsAuto functionality,
'             For a discussion see http://peltiertech.com/how-excel-calculates-automatic-chart-axis-limits/
'---------------------------------------------------------------------------------------
Private Function ChartMaxMin(DataMin As Double, DataMax As Double)
          Dim ChartMax As Double
          Dim ChartMin As Double
          Dim unitSize
1         On Error GoTo ErrHandler
2         If DataMin >= DataMax Then
3             ChartMin = DataMin - 1
4             ChartMax = DataMin + 1
5         Else
6             unitSize = Log((DataMax - DataMin) / 10) / Log(10#)
7             unitSize = Application.WorksheetFunction.Floor_Math(unitSize, 1)
8             unitSize = 10 ^ unitSize
9             ChartMax = DataMax + (DataMax - DataMin) / 20
10            ChartMax = Application.WorksheetFunction.Ceiling_Math(ChartMax, unitSize)
11            ChartMin = DataMin - (DataMax - DataMin) / 20
12            ChartMin = Application.WorksheetFunction.Floor_Math(ChartMin, unitSize)
13        End If
14        ChartMaxMin = sArrayStack(ChartMin, ChartMax)

15        Exit Function
ErrHandler:
16        Throw "#ChartMaxMin (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : EndOfNonZeroData
' Author    : Philip Swannell
' Date      : 29-Apr-2016
' Purpose   : Returns the contents of the first column for the maximum row such that all
'             elements on or after that row in columns 3 to 6 are zero.
'---------------------------------------------------------------------------------------
Function EndOfNonZeroData(TheData)
          Dim i As Long
          Dim ndg As Long
          Dim Res
1         On Error GoTo ErrHandler
2         Force2DArrayR TheData
3         ndg = LBound(TheData, 2) - 1
4         Res = TheData(UBound(TheData, 1), 1 + ndg)

5         For i = UBound(TheData, 1) - 1 To LBound(TheData, 1) + 1 Step -1
6             If TheData(i, 3 + ndg) = 0 And TheData(i, 4 + ndg) = 0 And TheData(i, 5 + ndg) = 0 And TheData(i, 6 + ndg) = 0 Then
7                 Res = TheData(i, 1 + ndg)
8             Else
9                 Exit For
10            End If
11        Next i
12        EndOfNonZeroData = Res
13        Exit Function
ErrHandler:
14        Throw "#EndOfNonZeroData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CleanOutDashboard
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Puts the dashboard sheet into a good state for releasing.
'---------------------------------------------------------------------------------------
Sub CleanOutDashboard()
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set SPH = CreateSheetProtectionHandler(shxVADashboard)
4         If IsInCollection(shxVADashboard.Names, "TheData") Then
5             RangeFromSheet(shxVADashboard, "TheData").Resize(, RangeFromSheet(shxVADashboard, "BottomHeaderRow").Columns.Count).Clear
6             shxVADashboard.Names("TheData").Delete
7         End If
8         AutoFitColumns RangeFromSheet(shxVADashboard, "BottomHeaderRow").Offset(-1).Resize(2), 2, sArrayRange(12, 6, 8)
          'get rid of the "extra columns" which appear in "DeveloperMode"
9         With shxVADashboard.Range("BottomHeaderRow")
10            sExpandDown(.Offset(-1, .Columns.Count).Resize(2, 3)).Clear
11            RemoveSortButtonsInRange .Offset(-2, .Columns.Count).Resize(1, 3)
12        End With
13        Exit Sub
ErrHandler:
14        Throw "#CleanOutDashboard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'Sub TestUpdatexVADashboard()
'1         On Error GoTo ErrHandler
'2         UpdatexVADashboard False, True, True, True, True, ConfigRange("AnalyticProvider").Value
'3         Exit Sub
'ErrHandler:
'4         SomethingWentWrong "#TestUpdatexVADashboard (line " & CStr(Erl) + "): " & Err.Description & "!"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdatexVADashboard
' Author    : Philip
' Date      : 04-Apr-2016
' Purpose   : Pastes data to the sheet xVADashboard and applies formatting.
'     Call stack is: UpdatexVADashboard (VBA) calls
'                    GetDataForDashboard (VBA) calls
'                    ConstructDashboard (R) which takes an argument PVOnlyMode
'             ConstructDashboard gets its "source data" from one of two places:
'             if PVOnlyMode is TRUE then data comes from gCachedDataForDashboardInPVOnlyMode,
'             which is generated by R method DataForPortfolioSheet.
'             If PVOnlyMode is FALSE the "source data" is the dataframe SCRiPTResults$PartyResults which
'             is created when the R function RunSCRiPT executes, as happens during execution of the VBA
'             method XVAFrontEndMain.
'---------------------------------------------------------------------------------------
Sub UpdatexVADashboard(ByVal PVOnlyMode As Boolean, DoPV As Boolean, DoCVA As Boolean, DoKVA As Boolean, PartitionByNetSet, Optional Numeraire As String, Optional OurName As String)
          Dim c As Range
          Dim DataToPaste As Variant
          Dim ExtraHeadersRange As Range
          Dim HeadersFromR As Variant
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim Target As Range
          Dim WithExtraCols As Boolean

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set SPH = CreateSheetProtectionHandler(shxVADashboard)
4         WithExtraCols = ConfigRange("DeveloperMode")

5         If IsInCollection(shxVADashboard.Names, "TheData") Then
6             RangeFromSheet(shxVADashboard, "TheData").Clear
7         End If
8         ResetSortButtons RangeFromSheet(shxVADashboard, "BottomHeaderRow").Rows(-1), False, False

9         If Numeraire <> "" Then
10            RangeFromSheet(shxVADashboard, "HeaderCell2").Value = "PV (" & Numeraire & ")"
11            RangeFromSheet(shxVADashboard, "HeaderCell3").Value = "PV (" & Numeraire & ")"
12        End If
13        If OurName <> "" Then
14            RangeFromSheet(shxVADashboard, "HeaderCell1").Value = "Banks Exposure to " + OurName
15        End If

16        If Not DoPV Or Not PartitionByNetSet Then
17            PVOnlyMode = True
18        End If

19        DataToPaste = GetDataForDashboard(PVOnlyMode, DoCVA, DoKVA, WithExtraCols, HeadersFromR)

20        Set Target = RangeFromSheet(shxVADashboard, "BottomHeaderRow").Cells(2, 1).Resize(sNRows(DataToPaste), sNCols(DataToPaste))
21        Target.Value = DataToPaste
22        With Target
23            .Parent.Names.Add "TheData", .Offset(0)
24            .NumberFormat = NF_Comma0dp
25            .HorizontalAlignment = xlHAlignCenter
26            AddGreyBorders .Offset(0)
27            .Columns(3).Resize(, 6).Interior.Color = Colour_LightGrey
28            .Columns(15).Resize(, 6).Interior.Color = Colour_LightGrey
29            .ColumnWidth = 20
30            AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 2, sArrayRange(12, 6, 8), 20
31            For Each c In .Columns(3).Cells
32                If VarType(c.Value) = vbString Then
33                    c.HorizontalAlignment = xlHAlignLeft
34                End If
35            Next
36        End With

37        With shxVADashboard.Range("BottomHeaderRow")
38            Set ExtraHeadersRange = .Offset(-1, .Columns.Count).Resize(2, 3)
39        End With
40        If WithExtraCols Then
41            With ExtraHeadersRange.Rows(1)
42                .MergeCells = True
43                .Value = "DeveloperMode Data"
44            End With
45            With ExtraHeadersRange.Rows(2)
46                .Value = sSubArray(HeadersFromR, 1, shxVADashboard.Range("BottomHeaderRow").Columns.Count + 1)
47            End With

48            With ExtraHeadersRange
49                .HorizontalAlignment = xlHAlignCenter
50                .VerticalAlignment = xlHAlignCenter
51                .Font.Bold = True
52                .Font.Color = RGB(255, 255, 255)
53                .Interior.Color = RGB(0, 102, 204)
54                AddGreyBorders .Offset(0)
55                RemoveSortButtonsInRange .Rows(0)
56                AddSortButtons .Rows(0), 2
57                AutoFitColumns .Rows(2).Resize(Target.Rows.Count + 1), 2, sArrayRange(12, 6, 8), 20
58            End With
59        Else
60            With ExtraHeadersRange
61                sExpandDown(.Offset(0)).Clear
62                RemoveSortButtonsInRange .Rows(0)
63            End With
64        End If

65        Exit Sub
ErrHandler:
66        Throw "#UpdatexVADashboard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function MyMax(a, b)
1         MyMax = IIf(a > b, a, b)
End Function
Function MyMin(a, b)
1         MyMin = IIf(a > b, b, a)
End Function

' ----------------------------------------------------------------
' Procedure Name: AmendDashboardForErrors
' Purpose:        The function GetDataForDashboard returns data for which PFEs etc will be zero when errors were encountered. This method looks at the "Status" flags in R to overwrite the Dashboard numbers when errors were encountered
' Procedure Kind: Function
' Procedure Access: Public
' Parameter DashBoardData ():
' Author: Philip Swannell
' Date: 02-06-2018
' ----------------------------------------------------------------
Function AmendDashboardForErrors(DashBoardData)
          Dim cn_PartyName
          Dim ColHeaders
          Dim DashboardParties
          Dim ErrorParties
          Dim ErrorString As String
          Dim i As Long
          Dim j As Long
          Dim NCDBD
          Dim NCPR As Long
          Dim PartyResults
          Dim ThisOneBad As Boolean

          'TODO PGS 18 Nov 2020 This function is no longer used. It needs to be modernised and used again!

1         On Error GoTo ErrHandler
2         If sIsErrorString(PartyResults) Then Throw "Unexpected error getting data SCRiPTResults$PartyResults from R - " + PartyResults
3         NCPR = sNCols(PartyResults)
4         NCDBD = sNCols(DashBoardData)

5         ColHeaders = sArrayTranspose(sSubArray(PartyResults, 1, 1, 1))
6         cn_PartyName = sMatch("PartyName", ColHeaders)
7         If sIsErrorString(cn_PartyName) Then Throw "Unexpected error. Cannot find column headed 'PartyName' in R expression SCRiPTResults$PartyResults"

8         ErrorParties = sSubArray(PartyResults, 2, cn_PartyName, , 1)
9         DashboardParties = sSubArray(DashBoardData, 1, 1, , 1)

10        For i = 1 To sNRows(DashBoardData)
11            ThisOneBad = False
12            ErrorString = ""
13            For j = NCPR To 1 Step -1
14                If Right(ColHeaders(j, 1), 6) = "Status" Then
15                    If sIsErrorString(PartyResults(i + 1, j)) Then
16                        ThisOneBad = True
17                        ErrorString = PartyResults(i + 1, j)
18                        Exit For
19                    End If
20                End If
21            Next j
22            If ThisOneBad Then
23                For j = 3 To NCDBD
24                    DashBoardData(i, j) = Empty
25                Next j
26                DashBoardData(i, 3) = "Error calculating results for '" & ErrorParties(i, 1) & "': " + ErrorString
27            End If
28        Next i

29        AmendDashboardForErrors = DashBoardData

30        Exit Function
ErrHandler:
31        Throw "#AmendDashboardForErrors (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetFromDictionary
' Author     : Philip Swannell
' Date       : 01-Dec-2021
' Purpose    : Access data from the dictionary created via call to JuliaEval("xva_main(..."). Vectors are translated to
'              one-column 2-d arrays though would be more efficient to call JuliaEval_LowLevel with argument
'              JuliaVectorToXLColumn set to true.
' Parameters :
'  D       :
'  key     :
'  FlipSign:
'  Default :
' -----------------------------------------------------------------------------------------------------------------------
Function GetFromDictionary(D As Dictionary, key As String, Optional FlipSign As Boolean, Optional Default, Optional Transpose As Boolean)
          Dim i As Long
          Dim item
          Dim j As Long
          Dim nudge As Long
          Dim NumDims As Long
          Dim NumElements As Long
          Dim Result()
          
1         On Error GoTo ErrHandler
2         If Not D.Exists(key) Then
3             If Not IsMissing(Default) Then
4                 If FlipSign Then
5                     GetFromDictionary = sArrayMultiply(-1, Default)
6                 Else
7                     GetFromDictionary = Default
8                 End If
9                 Exit Function
10            Else
11                Throw "Key '" + key + "' not found in Dictionary"
12            End If
13        End If

14        item = D(key)
15        NumDims = NumDimensions(item)
16        Select Case NumDims
              Case 0
17                Throw "item '" + key + "' in dictionary is a scalar when expecting a vector or array"
18            Case 1
                  'convert to 2-dimensional array, 1 column
19                NumElements = UBound(item) - LBound(item) + 1
20                nudge = LBound(item) - 1
21                ReDim Result(1 To NumElements, 1 To 1)
22                For i = 1 To NumElements
23                    Result(i, 1) = item(i + nudge)
24                    If FlipSign Then
25                        If IsNumber(Result(i, 1)) Then
26                            Result(i, 1) = -Result(i, 1)
27                        End If
28                    End If
29                Next i
30            Case 2
31                Result = item
32                If Transpose Then
33                    Result = sArrayTranspose(Result)
34                End If
35                If FlipSign Then
36                    For i = LBound(Result, 1) To UBound(Result, 1)
37                        For j = LBound(Result, 2) To UBound(Result, 2)
38                            If IsNumber(Result(i, j)) Then
39                                Result(i, j) = -Result(i, j)
40                            End If
41                        Next j
42                    Next i
43                End If
44                GetFromDictionary = Result
45        End Select

46        GetFromDictionary = Result
47        Exit Function
ErrHandler:
48        Throw "#GetFromDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConstructDashboardVBA
' Author     : Philip Swannell
' Date       : 19-Aug-2020
' Purpose    : Part of project to stop using BERT. Hence data for worksheet 'Dashboard' is derived from the Results.json
'              file rather than by by calling the R function ConstructDashboard that gets results from the SCRiPTResults R list object.
'              This method is a port to VBA of method ConstructDashboard in SCRiPTInterface.R
' Parameters :
'  Results        : A Dictionary - as returned by JuliaEval("xva_main(...)
'  PVOnlyMode     :
'  WhatIfPartyName:
'  WithExtraCols  :
' -----------------------------------------------------------------------------------------------------------------------
Function ConstructDashboardVBA(ByVal D As Dictionary, PVOnlyMode As Boolean, WhatIfPartyName As String, WithExtraCols)

          Dim CVA
          Dim CVAWhatIf
          Dim DVA
          Dim DVAWhatIf
          Dim FBA
          Dim FBAWhatIf
          Dim FCA
          Dim FCAWhatIf
          Dim FVAOld
          Dim FVAWhatIfOld
          Dim ImpactCVA
          Dim ImpactDVA
          Dim ImpactFBA
          Dim ImpactFCA
          Dim ImpactFVAOld
          Dim ImpactKVA
          Dim ImpactPV
          Dim ImpactxVA
          Dim KVA
          Dim KVAWhatIf
          Dim NA As Variant
          Dim NR As Long
          Dim PartyName
          Dim PV
          Dim PVWhatIf
          Dim TradeCount

1         On Error GoTo ErrHandler

2         If D.Exists("PartyResults") Then
3             Set D = D("PartyResults")
4         Else
5             Throw "Cannot find item 'PartyResults' in Dictionary Results"
6         End If

7         PV = GetFromDictionary(D, "PV", True)
8         PVWhatIf = GetFromDictionary(D, "PVWhatIf", True)
9         ImpactPV = sArraySubtract(PVWhatIf, PV)
10        PartyName = GetFromDictionary(D, "PartyName")
11        TradeCount = GetFromDictionary(D, "NumTrades")

12        NR = sNRows(PV)
13        If PVOnlyMode Then
14            NA = sReshape("", NR, 1)
15            CVA = NA: DVA = NA: FCA = NA: FBA = NA: KVA = NA
16            CVAWhatIf = NA: DVAWhatIf = NA: FCAWhatIf = NA
17            FBAWhatIf = NA: KVAWhatIf = NA: ImpactCVA = NA
18            ImpactDVA = NA: ImpactFCA = NA
19            ImpactFBA = NA: ImpactKVA = NA: ImpactxVA = NA
20        Else
              'Note the flip CVA <-> DVA below, necessary since trades are booked from the "Corporates" PoV, but we want CVA etc from the Banks' PoV
21            CVA = GetFromDictionary(D, "DVA", True)
22            DVA = GetFromDictionary(D, "CVA")
23            FCA = GetFromDictionary(D, "FCA_CP")
24            FBA = GetFromDictionary(D, "FBA_CP")
25            KVA = sReshape(0, NR, 1)
26            CVAWhatIf = GetFromDictionary(D, "DVAWhatIf", True)
27            DVAWhatIf = GetFromDictionary(D, "CVAWhatIf")
28            FCAWhatIf = GetFromDictionary(D, "FCAWhatIf_CP")
29            FBAWhatIf = GetFromDictionary(D, "FBAWhatIf_CP")
30            KVAWhatIf = sReshape(0, NR, 1)
31            ImpactCVA = sArraySubtract(CVAWhatIf, CVA)
32            ImpactDVA = sArraySubtract(DVAWhatIf, DVA)
33            ImpactFCA = sArraySubtract(FCAWhatIf, FCA)
34            ImpactFBA = sArraySubtract(FBAWhatIf, FBA)
35            ImpactKVA = sReshape(0, NR, 1)
36            ImpactxVA = sReshape(0, NR, 1)
37        End If

          Dim Data
          Dim LeftHeaders
          Dim TopHeaders

38        Data = sArrayRange(PartyName, TradeCount, PV, CVA, DVA, FCA, FBA, KVA, PVWhatIf, CVAWhatIf, DVAWhatIf, FCAWhatIf, FBAWhatIf, KVAWhatIf, ImpactPV, ImpactCVA, ImpactDVA, ImpactFCA, ImpactFBA, ImpactKVA, ImpactxVA)
39        TopHeaders = sArrayRange("PartyName", "TradeCount", "PV", "CVA", "DVA", "FCA", "FBA", "KVA", "PVWhatIf", "CVAWhatIf", "DVAWhatIf", "FCAWhatIf", "FBAWhatIf", "KVAWhatIf", "ImpactPV", "ImpactCVA", "ImpactDVA", "ImpactFCA", "ImpactFBA", "ImpactKVA", "ImpactxVA")
40        LeftHeaders = PartyName

41        If WithExtraCols Then
42            FVAOld = sArraySubtract(GetFromDictionary(D, "PV"), GetFromDictionary(D, "FundingPV"))  ' Old refers to calculation of FVA via changing discount factors, does not allow for split to FCA_CP and FBA_CP
43            FVAWhatIfOld = sArraySubtract(GetFromDictionary(D, "PVWhatIf"), GetFromDictionary(D, "FundingPVWhatIf"))
44            ImpactFVAOld = sArraySubtract(FVAWhatIfOld, FVAOld)
45            Data = sArrayRange(Data, FVAOld, FVAWhatIfOld, ImpactFVAOld)
46            TopHeaders = sArrayRange(TopHeaders, "FVAOld", "FVAWhatIfOld", "ImpactFVAOld")
47        End If

          'Now check the "Status" of calculations, replicates method AmendDashboardForErrors

          Dim AllOK
          Dim CVADVAStatus
          Dim ErrorString As String
          Dim FundingPVStatus
          Dim FundingPVWhatIfStatus
          Dim i As Long
          Dim j As Long
          Dim NCD As Long
          Dim PFEStatus
          Dim PVStatus
          Dim PVWhatIfStatus
48        AllOK = sReshape(0, NR, 1)
49        NCD = sNCols(Data)
50        PVStatus = GetFromDictionary(D, "PVStatus", False, AllOK)
51        FundingPVStatus = GetFromDictionary(D, "FundingPVStatus", False, AllOK)
52        PVWhatIfStatus = GetFromDictionary(D, "PVWhatIfStatus", False, AllOK)
53        FundingPVWhatIfStatus = GetFromDictionary(D, "FundingPVWhatIfStatus", False, AllOK)
54        CVADVAStatus = GetFromDictionary(D, "CVADVAStatus", False, AllOK)
55        PFEStatus = GetFromDictionary(D, "PFEStatus", False, AllOK)
          
56        For i = 1 To NR
57            ErrorString = ""
58            If sIsErrorString(PFEStatus(i, 1)) Then
59                ErrorString = PFEStatus(i, 1)
60            ElseIf sIsErrorString(CVADVAStatus(i, 1)) Then
61                ErrorString = CVADVAStatus(i, 1)
62            ElseIf sIsErrorString(FundingPVWhatIfStatus(i, 1)) Then
63                ErrorString = FundingPVWhatIfStatus(i, 1)
64            ElseIf sIsErrorString(PVWhatIfStatus(i, 1)) Then
65                ErrorString = PVWhatIfStatus(i, 1)
66            ElseIf sIsErrorString(FundingPVStatus(i, 1)) Then
67                ErrorString = FundingPVStatus(i, 1)
68            ElseIf sIsErrorString(PVStatus(i, 1)) Then
69                ErrorString = PVStatus(i, 1)
70            End If
71            If ErrorString <> "" Then
72                For j = 3 To NCD
73                    Data(i, j) = Empty
74                Next j
75                Data(i, 3) = "Error calculating results for '" & Data(i, 1) & "': " + ErrorString
76            End If
77        Next i

78        ConstructDashboardVBA = sArraySquare("", TopHeaders, LeftHeaders, Data)

79        Exit Function
ErrHandler:
80        Throw "#ConstructDashboardVBA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDataForDashboard
' Author    : Philip
' Date      : 04-Apr-2016
' Purpose   : Derive the data to put on the xVADashboard, makes a call to R function
'             ConstructDashboard and amends the data by looking at the HiddenSheet for data
'             placed there by method UpdateReturnOnRWAsByCounterparty.
'---------------------------------------------------------------------------------------
Function GetDataForDashboard(PVOnlyMode As Boolean, ByVal DoCVA As Boolean, ByVal DoKVA As Boolean, WithExtraCols As Boolean, ByRef HeadersFromR)
          Dim DataFromHiddenSheet As Variant
          Dim DataFromJulia As Variant
          Const ExpectedHeaders = "PartyName,TradeCount,PV,CVA,DVA,FCA,FBA,KVA,PVWhatIf,CVAWhatIf,DVAWhatIf,FCAWhatIf,FBAWhatIf,KVAWhatIf,ImpactPV,ImpactCVA,ImpactDVA,ImpactFCA,ImpactFBA,ImpactKVA,ImpactxVA"
          Const ExpectedExtraHeaders = ",FVAOld,FVAWhatIfOld,ImpactFVAOld"

          'Column numbers are correct once data coming back from R has its row and column headers stripped off
          Const cn_BankName = 1
          Const cn_CVA = 4
          Const cn_DVA = 5
          Const cn_FCA = 6
          Const cn_FBA = 7
          Const cn_KVA = 8
          Const cn_PVWhatIf = 9
          Const cn_CVAWhatIf = 10
          Const cn_DVAWhatIf = 11
          Const cn_FCAWhatIf = 12
          Const cn_FBAWhatIf = 13
          Const cn_KVAWhatIf = 14
          Const cn_ImpactPV = 15
          Const cn_ImpactCVA = 16
          Const cn_ImpactDVA = 17
          Const cn_ImpactFCA = 18
          Const cn_ImpactFBA = 19
          Const cn_ImpactKVA = 20
          Const cn_ImpactxVA = 21    ' We overwrite this column to include the impact of DVABenefit percentages and FVA charges percentages
          Dim BankNames As Variant
          Dim DVAbenefits As Variant
          Dim FVAcharges As Variant

          Dim N As Long

1         On Error GoTo ErrHandler

2     '    If PVOnlyMode Then
3     '        Throw "Method ConstructDashboardVBA not yet implemented for PVOnly mode"
4     '    Else
5             DataFromJulia = ConstructDashboardVBA(gResults, PVOnlyMode, gWHATIF, WithExtraCols)
6     '    End If

7         If DoKVA Or DoCVA Then
8             BankNames = sSubArray(DataFromJulia, 2, cn_BankName, , 1)
9             DVAbenefits = ThrowIfError(LookupBankInfo(BankNames, "DVA benefit %"))
10            FVAcharges = ThrowIfError(LookupBankInfo(BankNames, "FVA charge %"))
11            If sNRows(BankNames) = 1 Then
12                Force2DArray DVAbenefits
13                Force2DArray FVAcharges
14            End If
15        End If

16        DataFromHiddenSheet = sReshape("-", 3, 3)
17        HeadersFromR = sSubArray(DataFromJulia, 1, 2, 1)

18        DataFromJulia = sSubArray(DataFromJulia, 2, 2)
19        N = sNRows(DataFromJulia)
20        If sRowConcatenateStrings(HeadersFromR) <> (ExpectedHeaders & IIf(WithExtraCols, ExpectedExtraHeaders, "")) Then
21            Throw "Detected error in header row of data from call to R function ConstructDashboard"
22        End If

23        If DoCVA And DoKVA Then
              Dim i As Long
              Dim MatchIDs
24            MatchIDs = sMatch(sSubArray(DataFromJulia, 1, 1, , 1), sSubArray(DataFromHiddenSheet, 1, 1, , 1))

25            If N = 1 Then Force2DArray MatchIDs
26            For i = 1 To N
27                If IsNumber(MatchIDs(i, 1)) Then
28                    DataFromJulia(i, cn_KVA) = DataFromHiddenSheet(MatchIDs(i, 1), 2)
29                    DataFromJulia(i, cn_KVAWhatIf) = DataFromHiddenSheet(MatchIDs(i, 1), 3)
30                    If IsNumber(DataFromJulia(i, cn_KVA)) And IsNumber(DataFromJulia(i, cn_KVAWhatIf)) Then
31                        DataFromJulia(i, cn_ImpactKVA) = DataFromJulia(i, cn_KVAWhatIf) - DataFromJulia(i, cn_KVA)
32                    Else
33                        DataFromJulia(i, cn_ImpactKVA) = "#Error!"
34                    End If
35                    If IsNumber(DataFromJulia(i, cn_ImpactCVA)) And IsNumber(DataFromJulia(i, cn_ImpactDVA)) And IsNumber(DataFromJulia(i, cn_ImpactFCA)) And _
                          IsNumber(DataFromJulia(i, cn_ImpactFBA)) And IsNumber(DataFromJulia(i, cn_ImpactKVA)) Then
36                        If Not IsNumber(DVAbenefits(i, 1)) Then
37                            DataFromJulia(i, cn_ImpactxVA) = "Cannot get DVA benefit % from Lines workbook for bank " + BankNames(i, 1)
38                        ElseIf Not IsNumber(FVAcharges(i, 1)) Then
39                            DataFromJulia(i, cn_ImpactxVA) = "Cannot get FVA charge % from Lines workbook for bank " + BankNames(i, 1)
40                        Else
41                            DataFromJulia(i, cn_ImpactxVA) = DataFromJulia(i, cn_ImpactCVA) + _
                                  DataFromJulia(i, cn_ImpactDVA) * DVAbenefits(i, 1) + _
                                  DataFromJulia(i, cn_ImpactFBA) * FVAcharges(i, 1) + _
                                  DataFromJulia(i, cn_ImpactFCA) * FVAcharges(i, 1) + _
                                  DataFromJulia(i, cn_ImpactKVA)
42                        End If
43                    Else
44                        DataFromJulia(i, cn_ImpactxVA) = "#Error!"
45                    End If
46                Else
47                    DataFromJulia(i, cn_KVA) = "#Error!"
48                    DataFromJulia(i, cn_KVAWhatIf) = "#Error!"
49                    DataFromJulia(i, cn_ImpactKVA) = "#Error!"
50                End If
51            Next i
52        Else
53            For i = 1 To N
54                DataFromJulia(i, cn_KVA) = ""
55                DataFromJulia(i, cn_KVAWhatIf) = ""
56                DataFromJulia(i, cn_ImpactKVA) = ""
57                DataFromJulia(i, cn_ImpactxVA) = ""
58            Next i
59        End If

60        If DoCVA And Not DoKVA Then    'In this case we still show an xVA number even though KVA was not calculated because PFE was not calculated
61            For i = 1 To N
62                If IsNumber(DataFromJulia(i, cn_ImpactCVA)) And IsNumber(DataFromJulia(i, cn_ImpactDVA)) And IsNumber(DataFromJulia(i, cn_ImpactFCA)) And IsNumber(DataFromJulia(i, cn_ImpactFBA)) Then
63                    If Not IsNumber(DVAbenefits(i, 1)) Then
64                        DataFromJulia(i, cn_ImpactxVA) = "Cannot get DVA benefit % from Lines workbook for bank " + BankNames(i, 1)
65                    ElseIf Not IsNumber(FVAcharges(i, 1)) Then
66                        DataFromJulia(i, cn_ImpactxVA) = "Cannot get FVA charge % from Lines workbook for bank " + BankNames(i, 1)
67                    Else
68                        DataFromJulia(i, cn_ImpactxVA) = DataFromJulia(i, cn_ImpactCVA) + _
                              DataFromJulia(i, cn_ImpactDVA) * DVAbenefits(i, 1) + _
                              DataFromJulia(i, cn_ImpactFCA) * FVAcharges(i, 1) + _
                              DataFromJulia(i, cn_ImpactFBA) * FVAcharges(i, 1)
69                    End If
70                Else
71                    DataFromJulia(i, cn_ImpactxVA) = "#Error!"
72                End If
73            Next i
74        End If

75        If Not DoCVA Then
76            For i = 1 To N
77                DataFromJulia(i, cn_CVA) = ""
78                DataFromJulia(i, cn_DVA) = ""
79                DataFromJulia(i, cn_CVAWhatIf) = ""
80                DataFromJulia(i, cn_DVAWhatIf) = ""
81                DataFromJulia(i, cn_ImpactCVA) = ""
82                DataFromJulia(i, cn_ImpactDVA) = ""
83            Next i
84        End If

85      '  If PVOnlyMode Then
86      '      For i = 1 To N
87      '          DataFromJulia(i, cn_FCA) = ""
88      '          DataFromJulia(i, cn_FBA) = ""
89      '       '   DataFromJulia(i, cn_PVWhatIf) = ""
90      '          DataFromJulia(i, cn_FCAWhatIf) = ""
91      '          DataFromJulia(i, cn_FBAWhatIf) = ""
92      '          DataFromJulia(i, cn_ImpactFCA) = ""
93      '          DataFromJulia(i, cn_ImpactFBA) = ""
94      '        '  DataFromJulia(i, cn_ImpactPV) = ""
95      '      Next i
96      '  End If

97        GetDataForDashboard = DataFromJulia
98        Exit Function
ErrHandler:
99        GetDataForDashboard = "#GetDataForDashboard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub ClearCashflowDrilldownSheet()
          Dim b As Button
          Dim SPH As SolumAddin.clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shCashflowDrilldown)
3         shCashflowDrilldown.UsedRange.Clear
4         shCashflowDrilldown.UsedRange.EntireColumn.Delete

5         For Each b In shCashflowDrilldown.Buttons
6             b.Delete
7         Next
8         Exit Sub
ErrHandler:
9         Throw "#ClearCashflowDrilldownSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ViewTradeCashflows
' Author    : Philip Swannell
' Date      : 18-May-2016
' Purpose   : Pops up a sheet containing the trade and the flows of the trade.
'---------------------------------------------------------------------------------------
Function ViewTradeCashflows(Trade As Range)
          Dim b As Button
          Dim CashflowArray
          Dim CashflowsRange As Range
          Dim CellsBeneathButton As Range
          Dim CopyOfErr As String
          Dim HeaderText As String
          Dim i As Long
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TradeForR
          Dim TradeLabelsRange As Range
          Dim TradeValuesRange As Range
          Dim ValuationFunction As String
          Dim WasProtected As Boolean
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         Throw "Trade Cashflows not implemented (yet?) in Julia"

3         If Not ModelExists() Then Throw "Please rebuild model before viewing trade cashflows", True
4         ValuationFunction = Trade(1, gCN_TradeType)

5         TradeForR = PortfolioTradesToJuliaTrades(Trade.Value2, True, False)
6         SaveDataframe TradeForR, "input_onetrade"
7         CashflowArray = sSubArray(CashflowArray, 1, 2)
8         If ValuationFunction = "InterestRateSwap" Or ValuationFunction = "CrossCurrencySwap" Or ValuationFunction = "CapFloor" Or ValuationFunction = "InflationZCSwap" Then
9             CashflowArray = sSuppressNAs(CashflowArray)
10        End If

11        Set SUH = CreateScreenUpdateHandler()

12        Application.DisplayAlerts = False

13        Set ws = shCashflowDrilldown
14        WasProtected = ThisWorkbook.ProtectStructure
15        If WasProtected Then ThisWorkbook.Protect , False
16        ws.Visible = xlSheetVisible
17        ClearCashflowDrilldownSheet
18        Set SPH = CreateSheetProtectionHandler(ws)

19        ws.Cells(1, 1).ColumnWidth = 2

20        HeaderText = "Cashflows for trade " + CStr(Trade.Cells(1, gCN_TradeID).Value)
21        With ws.Cells(1, 2)
22            .Value = HeaderText
23            .Font.Size = 22
24        End With

25        Set CashflowsRange = ws.Cells(4, 5).Resize(sNRows(CashflowArray), sNCols(CashflowArray))

26        CashflowsRange.Value = CashflowArray

27        ApplySolumFormatting CashflowsRange, "Cashflows and PVs from " + CStr(ConfigRange("OurName").Value) + "'s perspective.", False

28        Set TradeValuesRange = ws.Cells(4, 3).Resize(Trade.Columns.Count)
29        TradeValuesRange.Value2 = sArrayTranspose(Trade.Value2)
30        Set TradeLabelsRange = TradeValuesRange.Columns(0)
31        shHiddenSheet.Calculate
32        TradeLabelsRange.Value = sArrayTranspose(RangeFromSheet(shHiddenSheet, "SingleRowHeaders").Value)
33        For i = 1 To Trade.Columns.Count
34            TradeValuesRange.Cells(i, 1).NumberFormat = Trade(1, i).NumberFormat
35            TradeValuesRange.Cells(i, 1).Interior.Color = Trade(1, i).Interior.Color
36            TradeValuesRange.Cells(i, 1).Font.Color = Trade(1, i).Font.Color
37        Next i
38        With TradeLabelsRange
39            .Interior.Color = RGB(0, 102, 204)
40            .Font.Color = RGB(255, 255, 255)
41        End With
42        With Application.Union(TradeLabelsRange, TradeValuesRange)
43            .HorizontalAlignment = xlHAlignLeft
44            AutoFitColumns .Offset(0), 0.5, , 20
45            AddGreyBorders .Offset(0)
46        End With

47        Set CellsBeneathButton = TradeLabelsRange.Cells(TradeLabelsRange.Rows.Count + 2, 1).Resize(2)
48        With CellsBeneathButton
49            Set b = ws.Buttons.Add(.Left, .Top, .Width, .Height)
50        End With
51        With b
52            .Caption = "OK"
53            .OnAction = "'" & ThisWorkbook.Name & "'" & "!HideCashflowsDrilldownsheet"
54            .Font.Name = "Calibri"
55            .Font.Size = 11
56            .Font.Bold = True
57        End With

58        Application.GoTo TradeValuesRange.Cells(1, 1)
59        Application.OnKey "{ESCAPE}", "'" & ThisWorkbook.Name & "'" & "!HideCashflowsDrilldownsheet"

60        Exit Function
ErrHandler:
61        CopyOfErr = "#ViewTradeCashflows (line " & CStr(Erl) + "): " & Err.Description & "!"
62        Throw CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : HideCashflowsDrilldownsheet
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : The CashflowDrilldown sheet has an "OK" button that calls this method
'---------------------------------------------------------------------------------------
Sub HideCashflowsDrilldownsheet()
1         On Error GoTo ErrHandler
          Dim WasProtected As Boolean
2         WasProtected = ThisWorkbook.ProtectStructure
3         If WasProtected Then ThisWorkbook.Protect , False
4         shPortfolio.Activate
5         shCashflowDrilldown.Visible = xlSheetHidden

6         If WasProtected Then ThisWorkbook.Protect , True
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#HideCashflowsDrilldownsheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub TestAP_Results()
          Dim ControlString As String
          Dim ID As String
          Dim IDIsNetSet As Boolean
          Dim Res
          Dim WithHeaders As Boolean

1         If gResults Is Nothing Then
2             ReloadResults
3         End If

4         ID = "T000066"
5         ID = "BARC_GB_LON"
6         ControlString = "Date,Time,OurSP,TheirSP,OurEE,TheirEE,OurEPE,TheirEPE,OurENE,TheirENE,OurPaths,TheirPaths"
7         WithHeaders = True
8         IDIsNetSet = True

          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
9         t1 = sElapsedTime
10        Res = GetResultsFromJulia(ID, ControlString, WithHeaders, IDIsNetSet)
11        t2 = sElapsedTime
13        t3 = sElapsedTime
14        Debug.Print t2 - t1, t3 - t2

15        g Res

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetResultsFromJulia
' Author     : Philip Swannell
' Date       : 24-Aug-2020
' Purpose    : Port to VBA of R function SCRiPT_Results (that was wrapped by VBA function of the same name, defined in SolumSCRiPTUtils.xlam). Takes data from Dictionary gResults
' Parameters :
'  ID           :
'  ControlString:
'  WithHeaders  :
'  IDIsNetSet   :
' -----------------------------------------------------------------------------------------------------------------------
Function GetResultsFromJulia(ID As String, ControlString As String, Optional WithHeaders As Boolean = False, Optional IDIsNetSet As Boolean)
          Dim ControlArray
          Dim Data As Dictionary
          Dim i As Long
          Dim KeyName As String
          Const ErrControlString = "ControlString must be comma delimited. Each token must be: 'Time', 'Date' " & _
              "or ABC where A is 'Our' or 'Their'; B is 'SP', 'EE', 'EPE', 'ENE', 'PFE' or 'Paths'; and C is absent or 'WhatIf'"
          Dim AnchorDate
          Dim Atrib As String
          Dim ChangeSign As Boolean
          Dim DoTranspose As Boolean
          Dim Status As String
          Dim ThisCol
          Dim ThisControl As String
          Dim ThisHeader

1         On Error GoTo ErrHandler
2         ControlArray = sTokeniseString(UCase(ControlString))
3         If gResults Is Nothing Then Throw "Data has not yet been generated"
4         KeyName = IIf(IDIsNetSet, "PartyExposures", "TradeExposures")
5         If gResults.Exists(KeyName) Then
6             Set Data = gResults(KeyName)
7         Else
8             Throw "Cannot find Dictionary element '" + KeyName + "'"
9         End If
10        If Not Data.Exists(ID) Then Throw IIf(IDIsNetSet, "Counterparty '", "Trade '") & ID & "' not found"
11        Set Data = Data(ID)

          Dim rx As New VBScript_RegExp_55.RegExp
12        With rx
13            .IgnoreCase = False
14            .Pattern = "^(OUR|THEIR)(EE|EPE|ENE|PFE|PATHS)(WHATIF|)$"
15            .Global = False
16        End With

17        Status = Data("PFEStatus")
18        If Status <> "OK" Then Throw "There was an error when the data was generated for '" & ID & "':" & Status

          Dim STK As clsHStacker
19        Set STK = CreateHStacker()

20        For i = 1 To sNRows(ControlArray)
21            ThisControl = ControlArray(i, 1)
22            Select Case ThisControl
                  Case "OURSP"
23                    ThisHeader = "Self Surv Prob"
24                    ThisCol = GetFromDictionary(Data, "SP_Self")
25                Case "THEIRSP"
26                    ThisHeader = "CP Surv Prob"
27                    ThisCol = GetFromDictionary(Data, "SP_CP")
28                Case "TIME"
29                    ThisCol = GetFromDictionary(Data, "Time")
30                    ThisHeader = "Time"
31                Case "DATE"
32                    AnchorDate = Data("AnchorDate")
33                    ThisCol = GetFromDictionary(Data, "Time")
34                    ThisCol = sArrayMultiply(ThisCol, 365.25) 'In line with R function timetoexceldate
35                    ThisCol = sArrayAdd(ThisCol, AnchorDate)
36                    ThisHeader = "Date"
37                Case Else
38                    If Not rx.Test(ThisControl) Then Throw ErrControlString
39                    If Not IDIsNetSet Then
40                        If InStr(ControlArray(i, 1), "WHATIF") > 0 Then
41                            Throw "ControlString tokens cannot include 'WhatIf' when ID is a TradeID"
42                        End If
43                    End If
44                    If InStr(ThisControl, "OURPFE") > 0 Then Throw "OurPFE is not available" ''TODO - calculate on-the-fly from the Paths?
45                    ThisHeader = Replace(Replace(Replace(Replace(ThisControl, "OUR", "Our "), "THEIR", "Their "), "PATHS", "Path"), "WHATIF", "WhatIf ")
46                    Atrib = Replace(Replace(Replace(Replace(ThisControl, "OUR", ""), "THEIR", ""), "PATHS", "Paths"), "WHATIF", "WhatIf")
                      'To save memory if there are no WhatIf trades we don't save data for the "With WhatIf" Exposures, so we need to look at the "regular" Exposures
47                    If IDIsNetSet Then
48                        If Data("NumWhatIfTrades") = 0 Then
49                            Atrib = Replace(Atrib, "WhatIf", "")
50                        End If
51                    End If
52                    If UCase(Left(ThisControl, 3)) = "OUR" Then
53                        ChangeSign = False
54                    Else
55                        ChangeSign = True
                          'THEIR' case
56                        If InStr(Atrib, "EPE") > 0 Then
57                            Atrib = Replace(Atrib, "EPE", "ENE")
58                        ElseIf InStr(Atrib, "ENE") > 0 Then
59                            Atrib = Replace(Atrib, "ENE", "EPE")
60                        End If
61                    End If

62                    DoTranspose = False
63                    If InStr(Atrib, "Paths") > 0 Then
64                        DoTranspose = True
65                        If Not IsArray(Data(Atrib)) Then
66                            Throw "Paths not available. " + gProjectName + " workbook Config sheet must have SavePaths set to TRUE when the data is generated"
67                        End If
68                    End If

69                    ThisCol = GetFromDictionary(Data, Atrib, ChangeSign, , DoTranspose)
                      
70                    If sNCols(ThisCol) > 1 Then ThisHeader = sArrayRange(ThisHeader, sReshape("", 1, sNCols(ThisCol) - 1))

71            End Select

72            If WithHeaders Then
73                STK.Stack2D sArrayStack(ThisHeader, ThisCol)
74            Else
75                STK.Stack2D ThisCol
76            End If

77        Next i

78        GetResultsFromJulia = STK.report

79        Exit Function
ErrHandler:
80        GetResultsFromJulia = "#GetResultsFromJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub TestTradeDataForCounterpartyViewerSheet()

1         On Error GoTo ErrHandler

2         ReloadResults

3         g TradeDataForCounterpartyViewerSheet("BARC_GB_LON", True, gWHATIF)

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestTradeDataForCounterpartyViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TradeDataForCounterpartyViewerSheet
' Author     : Philip Swannell
' Date       : 26-Aug-2020
' Purpose    : Port to VBA of R function Data2ForCounterpartyViewerSheet, but currently only used
' Parameters :
'  Counterparty       :
'  IncludeHyptheticals:
'  WhatIfPartyName    :
' -----------------------------------------------------------------------------------------------------------------------
Function TradeDataForCounterpartyViewerSheet(Counterparty As String, IncludeHyptheticals As Boolean, WhatIfPartyName As String)
1         On Error GoTo ErrHandler
          Dim AllTrades
          Dim ChooseVector As Variant
          Dim i As Long
          Dim j As Long
          Dim TD As Dictionary
          Dim TheseTrades
          Dim TRD As Dictionary
          
2         If gResults Is Nothing Then Throw "Data has not yet been generated"
3         Set TD = gResults("Trades")

4         AllTrades = HStack(TD("TradeID"), TD("ValuationFunction"), TD("Counterparty"), TD("StartDate"), TD("EndDate"))
5         AllTrades = VStack(HStack("TradeID", "ValuationFunction", "Counterparty", "StartDate", "EndDate"), AllTrades)
          
6         ChooseVector = sReshape(False, sNRows(AllTrades), 1)
7         ChooseVector(1, 1) = True
8         For i = 2 To sNRows(AllTrades)
9             If AllTrades(i, 3) = Counterparty Then
10                ChooseVector(i, 1) = True
11            ElseIf AllTrades(i, 3) = WhatIfPartyName Then
12                ChooseVector(i, 1) = IncludeHyptheticals
13            End If
14        Next i
15        TheseTrades = sMChoose(AllTrades, ChooseVector)
       
          Dim AllTradeResults
          Dim MatchIDs
          Dim Result
          Dim TheseTradeResults
          
16        Set TRD = gResults("TradeResults")
          
17        AllTradeResults = HStack(TRD("TradeID"), TRD("PV"), TRD("DVA"), TRD("CVA"), TRD("FCA_CP"), TRD("FBA_CP"), TRD("FundingPV"))
18        AllTradeResults = VStack(HStack("TradeID", "PV", "DVA", "CVA", "FCA_CP", "FBA_CP", "FundingPV"), AllTradeResults)

19        AllTradeResults(1, 3) = "CVA" 'Code below also flips the sign of the contents of these columns...
20        AllTradeResults(1, 3) = "DVA"
          
21        MatchIDs = sMatch(sSubArray(TheseTrades, 1, 1, , 1), sSubArray(AllTradeResults, 1, 1, , 1))
22        TheseTradeResults = sIndex(AllTradeResults, MatchIDs)
          
23        For i = 2 To sNRows(TheseTradeResults)
24            For j = 4 To 5
25                If VarType(TheseTrades(i, j)) <> vbDate Then
27                        TheseTrades(i, j) = CVErr(xlErrNA)
33                End If
34            Next j

              'Flip PV
35            TheseTradeResults(i, 2) = -TheseTradeResults(i, 2)
              
36            TheseTradeResults(i, 3) = -TheseTradeResults(i, 3)
37            TheseTradeResults(i, 4) = -TheseTradeResults(i, 4)
              'Flip sign of FundingPV
38            TheseTradeResults(i, 7) = -TheseTradeResults(i, 7)
              'Overwrite FundingPV with FVA = FundingPV - PV
39            TheseTradeResults(i, 7) = TheseTradeResults(i, 7) - TheseTradeResults(i, 2)
          
40        Next i
          'relabel column
41        TheseTradeResults(1, 7) = "FVA"
42        Result = sArrayRange(TheseTrades, sSubArray(TheseTradeResults, 1, 2))
         
43        TradeDataForCounterpartyViewerSheet = Result
          
44        Exit Function
ErrHandler:
45        Throw "#TradeDataForCounterpartyViewerSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


