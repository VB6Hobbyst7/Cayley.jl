Attribute VB_Name = "modTradesViewer"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modTradesViewer
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Code relating to the TradesViewer sheet
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private Const m_HeaderHW = "PV (Hull-White)"
Private Const m_HeaderMTM = "Airbus MTM"
Private m_HeaderDiff

Public gBlockChangeEvent As Boolean

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearOutTradeValuesSheet
' Author    : Philip Swannell
' Date      : 05-Oct-2016
' Purpose   : Removes all date from the TradeValues sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearOutTradeValuesSheet()
          Dim Res As Variant
          Dim SPH As clsSheetProtectionHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set ws = shTradesViewer
3         Set SPH = CreateSheetProtectionHandler(ws)
          Dim b As Button
          Dim N As Name
4         For Each b In ws.Buttons
5             If Len(b.Caption) <= 1 Then
6                 b.Delete
7             End If
8         Next b
9         For Each N In ws.Names
10            N.Delete
11        Next N
12        ws.UsedRange.EntireColumn.Delete
13        ws.UsedRange.EntireRow.Delete
14        With ws.Cells(1, 1)
15            .Value = "Trades Viewer"
16            .Font.Size = 22
17        End With

18        Res = ws.UsedRange.Rows.Count        'Has the effect of resetting the UsedRange to top-left cell
19        Res = ws.UsedRange.Columns.Count
20        Exit Sub
ErrHandler:
21        Throw "#ClearOutTradeValuesSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowHelpOnTradesViewer
' Author    : Philip Swannell
' Date      : 07-Oct-2016
' Purpose   : Show a description of the TradeList sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowHelpOnTradesViewer()
          Dim Msg As String
1         On Error GoTo ErrHandler
              
2         Msg = "The TradesViewer sheet shows trades taken from the trades workbook, in " & _
              "the form passed to the Julia code." & vbLf & vbLf & _
              "You can choose to show:" & vbLf & vbLf & _
              "Filtered Trades" & vbLf & _
              "The trades in the trades workbook that match the filters on the CreditUsage sheet." & vbLf & vbLf & _
              "All Trades" & vbLf & _
              "All the trades in the trades workbook." & vbLf & vbLf & _
              "Extra Trades" & vbLf & _
              "The trades that correspond to the current values for ""ExtraTradesAre"" and" & _
              " ""Amounts USD"" on the CreditUsage sheet." & vbLf & vbLf & _
              "PVs calculated using the Hull-White model implemented in Julia are displayed and " & _
              "compared with a PV read from the 'MTM' column of the trades workbook." & vbLf & vbLf & _
              "PVs are shown from the banks' perspective, i.e. the negative of the PV from Airbus's perspective."

3         MsgBoxPlus Msg, vbInformation, "TradesViewer"
4         Exit Sub
ErrHandler:
5         Throw "#ShowHelpOnTradesViewer (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetCellComment
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : Adds a comment to a cell and makes it appear in Calibri 11. Comment must be
'             passed including line feed characters
' -----------------------------------------------------------------------------------------------------------------------
Function SetCellComment(c As Range, ByVal Comment As String, InsertBreaks As Boolean)
1         On Error GoTo ErrHandler

2         If InsertBreaks Then
3             Comment = sConcatenateStrings(sJustifyText(Comment, "Calibri", 11, 300), vbLf)
4         End If

5         c.ClearComments
6         c.AddComment
7         c.Comment.Visible = False
8         c.Comment.Text Text:=Comment
9         With c.Comment.Shape.TextFrame
10            .Characters.Font.Name = "Calibri"
11            .Characters.Font.Size = 11
              'Do not set .AutoSize to True - it's very slow. See Robin De Schepper's answer at
              'https://stackoverflow.com/questions/28670030/cell-comments-modifying-commands-are-extremely-slow-in-excel-vba
              ' .AutoSize = True
12        End With
13        Exit Function
ErrHandler:
14        Throw "#SetCellComment (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuTradesViewerSheet
' Author    : Philip Swannell
' Date      : 06-Oct-2016
' Purpose   : Attached to "Menu..." button on sheet TradesViewer
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuTradesViewerSheet()
          Const chHelp = "&About this sheet"
          Const FidHelp = 49
          Const chRefresh = "&Refresh..."
          Const FidRefresh = 1987
          Dim Choice
          Dim Choices
          Dim FIDs

1         On Error GoTo ErrHandler
2         RunThisAtTopOfCallStack
          
3         JuliaLaunchForCayley

4         Choices = sArrayStack(chRefresh, "--" & chHelp)
5         FIDs = sArrayStack(FidRefresh, FidHelp)
6         Choice = ShowCommandBarPopup(Choices, FIDs, , , ChooseAnchorObject())
7         Select Case Choice
              Case "#Cancel!"
                  'Nothing to do
8             Case Unembellish(chRefresh)
9                 MenuCreditUsageSheet "ShowTrades"
10            Case Unembellish(chHelp)
11                ShowHelpOnTradesViewer
12            Case Else
13                Throw "Unrecognised choice in menu: " & CStr(Choice)
14        End Select
15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#MenuTradesViewerSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MatchThrow
' Author    : Philip Swannell
' Date      : 06-Oct-2016
' Purpose   : wrapper to sMatch but with more informative error
' -----------------------------------------------------------------------------------------------------------------------
Private Function MatchThrow(LookupValue As String, LookupArray As Variant, Optional ThrowOnError As Boolean = True)
          Dim MatchRes As Variant
1         On Error GoTo ErrHandler
2         MatchRes = sMatch(LookupValue, LookupArray)
3         If IsNumber(MatchRes) Then
4             MatchThrow = MatchRes
5         Else
6             If ThrowOnError Then
7                 Throw "Cannot find '" & LookupValue & "' in headers of trade data"
8             Else
9                 MatchThrow = LookupValue & " not found"
10            End If
11        End If
12        Exit Function
ErrHandler:
13        Throw "#MatchThrow (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : EURNotional
' Author    : Philip Swannell
' Date      : 06-Oct-2016
' Purpose   : For an array of trades in Julia format we want the EUR Notional of each trade so
'             that we can express the difference between "Their" PV and "our "PV" in basis
'             points of Notional.
' -----------------------------------------------------------------------------------------------------------------------
Function EURNotional(Trades, ModelBareBones As Dictionary)
          Dim FxSpot As Double
          Dim Headers
          Dim i As Long
          Dim Notional() As Variant
          Dim NotionalColumn As Long
          Dim NumTrades As Long
          Dim ValuationFunction As String

          Dim cn_Currency As Long
          Dim cn_Notional As Long
          Dim cn_PayAmortNotionals As Long
          Dim cn_PayCurrency As Long
          Dim cn_PayNotional As Long
          Dim cn_ReceiveAmortNotionals As Long
          Dim cn_ReceiveCurrency As Long
          Dim cn_ReceiveNotional As Long
          Dim cn_ValuationFunction As Long

1         On Error GoTo ErrHandler

2         Headers = sArrayTranspose(sSubArray(Trades, 1, 1, 1))

3         cn_ValuationFunction = MatchThrow("ValuationFunction", Headers)
4         cn_Currency = MatchThrow("Currency", Headers)
5         cn_PayCurrency = MatchThrow("PayCurrency", Headers)
6         cn_ReceiveCurrency = MatchThrow("ReceiveCurrency", Headers)
7         cn_Notional = MatchThrow("Notional", Headers)
8         cn_PayNotional = MatchThrow("PayNotional", Headers)
9         cn_ReceiveNotional = MatchThrow("ReceiveNotional", Headers)
10        cn_PayAmortNotionals = MatchThrow("PayAmortNotionals", Headers)
11        cn_ReceiveAmortNotionals = MatchThrow("ReceiveAmortNotionals", Headers)

12        NumTrades = sNRows(Trades) - 1
13        ReDim Notional(1 To NumTrades)

14        For i = 1 To NumTrades
15            ValuationFunction = Trades(i + 1, cn_ValuationFunction)
16            Select Case ValuationFunction
                  Case "FxForward"
17                    If Trades(i + 1, cn_ReceiveCurrency) = "EUR" Then
18                        Notional(i) = Trades(i + 1, cn_ReceiveNotional)
19                    ElseIf Trades(i + 1, cn_PayCurrency) = "EUR" Then
20                        Notional(i) = Trades(i + 1, cn_PayNotional)
21                    Else
22                        FxSpot = 0
23                        On Error Resume Next
24                        FxSpot = MyFxPerBaseCcy(Trades(i + 1, cn_ReceiveCurrency), "EUR", ModelBareBones)
25                        On Error GoTo ErrHandler
26                        If FxSpot = 0 Then

27                            Notional(i) = "Cannot find Fx Rates for " & Trades(i + 1, cn_ReceiveCurrency)
28                        Else
29                            Notional(i) = Trades(i + 1, cn_ReceiveNotional) * FxSpot
30                        End If
31                    End If

32                Case "FxOption"
33                    If Trades(i + 1, cn_Currency) = "EUR" Then
34                        Notional(i) = Trades(i + 1, cn_Notional)
35                    Else
36                        FxSpot = 0
37                        On Error Resume Next
38                        FxSpot = MyFxPerBaseCcy(Trades(i + 1, cn_Currency), "EUR", ModelBareBones)
39                        On Error GoTo ErrHandler
40                        If FxSpot = 0 Then
41                            Notional(i) = "Cannot find Fx Rates for " & Trades(i + 1, cn_Currency)
42                        Else
43                            Notional(i) = Trades(i + 1, cn_Notional) * FxSpot
44                        End If
45                    End If

46                Case "InterestRateSwap", "CrossCurrencySwap"
47                    If Trades(i + 1, cn_ReceiveNotional) <> 0 Then
48                        NotionalColumn = cn_ReceiveNotional
49                    Else
50                        NotionalColumn = cn_ReceiveAmortNotionals
51                    End If
52                    If Trades(i + 1, cn_ReceiveCurrency) = "EUR" Then
53                        Notional(i) = CDbl(sTokeniseString(CStr(Trades(i + 1, NotionalColumn)), ";")(1, 1))
54                    Else
55                        FxSpot = 0
56                        On Error Resume Next
57                        FxSpot = MyFxPerBaseCcy(Trades(i + 1, cn_ReceiveCurrency), "EUR", ModelBareBones)
58                        On Error GoTo ErrHandler
59                        If FxSpot = 0 Then
60                            Notional(i) = "Cannot find Fx Rates for " & Trades(i + 1, cn_ReceiveCurrency)
61                        Else
62                            Notional(i) = CDbl(sTokeniseString(CStr(Trades(i + 1, NotionalColumn)), ";")(1, 1)) * FxSpot
63                        End If
64                    End If
65            End Select
66        Next i

67        EURNotional = sArrayTranspose(Notional)
68        Exit Function
ErrHandler:
69        Throw "#EURNotional (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowTrades
' Author    : Hermione Glyn, Philip Swannell
' Date      : 21-Sep-2016
' Purpose   : Displays all trades or trades matching the current filters in the sheet
'             TradeValuesHW together with their PVs in EUR
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowTrades()

          Dim BaseModelName As String
          Dim c As Range
          Dim CheckBoxText As String
          Dim CurrenciesToInclude As String
          Dim EURNotionals As Variant
          Dim Filter1Value As Variant
          Dim Filter2Value As Variant
          Dim FilterBy1 As Variant
          Dim FilterBy2 As Variant
          Dim FlipTrades As Boolean
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim HelpMethodName As String
          Dim IncludeAssetClasses As String
          Dim IncludeFutureTrades As Boolean
          Dim Mwb As Workbook
          Dim Numeraire As String
          Dim NumTrades As Long
          Dim NumValueColumns As Long
          Dim PortfolioAgeing As Double
          Dim ProductCreditLimits As String
          Dim ShockedModelName As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim Target As Range
          Dim TC As TradeCount
          Dim TradesJuliaFormat As Variant
          Dim TradesScaleFactor As Double
          Dim twb As Workbook
          Dim UseHistoricalFxVol As Boolean
          Dim ValueDiff As Variant
          Dim ValuesCalypso As Variant
          Dim WithFxTrades As Boolean
          Dim WithRatesTrades As Boolean
          Dim ws As Worksheet
          Const InsertValuesAfterCol = 5
          Const chAll = "All Trades"
          Const chFiltered = "Filtered Trades"
          Const chExtraTrades = "Extra Trades"
          Dim TopText As String
          Static Choice As String        'Make static so that dialog remembers previous choice
          Static WithValues As Boolean        'Make static so that dialog remembers previous choice
          Dim AnchorDate As Date
          Dim oldBlockChange As Boolean
          Dim ValueColumns
          Dim ValueHeaders

1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent

3         gBlockChangeEvent = True

4         CheckBoxText = "Calculate PVs (Hull-White)"

5         TopText = "What trades do you want to show?"
6         HelpMethodName = "'" & ThisWorkbook.Name & "'" & "!ShowHelpOnTradesViewer"
7         Choice = ShowOptionButtonDialog(sArrayStack(chFiltered, chAll, chExtraTrades), "Trades Viewer", TopText, Choice, , , CheckBoxText, WithValues, HelpMethodName)
8         Select Case Choice
              Case chAll, chFiltered, chExtraTrades
9             Case Else
10                GoTo EarlyExit
11        End Select

12        g_StartRunCreditUsageSheet = sElapsedTime()

13        Set SUH = CreateScreenUpdateHandler()
14        Set Mwb = OpenMarketWorkbook(True, False)
15        Set twb = OpenTradesWorkbook(True, False)

          Dim ModelBareBones As Dictionary

16        FxShock = RangeFromSheet(shCreditUsage, "FxShock", True, False, False, False, False).Value
17        FxVolShock = RangeFromSheet(shCreditUsage, "FxVolShock", True, False, False, False, False).Value

18        If WithValues Or Choice = chExtraTrades Then
19            JuliaLaunchForCayley
20            BuildModelsInJulia False, FxShock, FxVolShock
21            Set ModelBareBones = gModel_CM
22            AnchorDate = ModelBareBones("AnchorDate")
23        Else
24            AnchorDate = Date
25        End If

26        Numeraire = NumeraireFromMDWB()

27        If Choice = chExtraTrades Then
28            UseHistoricalFxVol = False
29            PortfolioAgeing = 0
30            BaseModelName = MN_CM
31            ShockedModelName = MN_CMS

32            TradesJuliaFormat = ConstructExtraTrades(RangeFromSheet(shCreditUsage, "ExtraTradesAre"), _
                  ModelBareBones, _
                  RangeFromSheet(shCreditUsage, "ExtraTradeAmounts"), _
                  RangeFromSheet(shCreditUsage, "ExtraTradeLabels"), False)
                  
33        Else
34            FlipTrades = True
35            If Choice = chAll Then
36                FilterBy1 = "None"
37                Filter1Value = "None"
38                FilterBy2 = "None"
39                Filter2Value = "None"
40                IncludeFutureTrades = RangeFromSheet(shCreditUsage, "IncludeFutureTrades", False, False, True, False, False).Value2
41                IncludeAssetClasses = "Rates and Fx"
42                ProductCreditLimits = "Global Calc"
43                PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False).Value2

44                TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value2
45                CurrenciesToInclude = "All"
46                UseHistoricalFxVol = False
47            ElseIf Choice = chFiltered Then
48                FilterBy1 = RangeFromSheet(shCreditUsage, "FilterBy1", False, True, False, False, False).Value2
49                Filter1Value = RangeFromSheet(shCreditUsage, "Filter1Value", True, True, True, False, False).Value2
50                FilterBy2 = RangeFromSheet(shCreditUsage, "FilterBy2", False, True, False, False, False).Value2
51                Filter2Value = RangeFromSheet(shCreditUsage, "Filter2Value", True, True, True, False, False).Value2
52                IncludeFutureTrades = RangeFromSheet(shCreditUsage, "IncludeFutureTrades", False, False, True, False, False).Value2
53                IncludeAssetClasses = RangeFromSheet(shCreditUsage, "IncludeAssetClasses", False, True, False, False, False).Value2
54                ProductCreditLimits = FirstElement(LookupCounterpartyInfo(Filter1Value, "Product Credit Limits", "Global Calculation", "Global Calculation"))
55                PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False).Value2
56                TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value2
57                CurrenciesToInclude = RangeFromSheet(shConfig, "CurrenciesToInclude", False, True, False, True, False)
58                UseHistoricalFxVol = FirstElementOf(LookupCounterpartyInfo(Filter1Value, "Volatility Input", , "MARKET IMPLIED")) = "HISTORICAL"
59            End If

60            SetBooleans IncludeAssetClasses, ProductCreditLimits, "Foo", False, WithFxTrades, WithRatesTrades, False

61            If UseHistoricalFxVol Then
62                BaseModelName = MN_CMH
63                ShockedModelName = MN_CMHS
64            Else
65                BaseModelName = MN_CM
66                ShockedModelName = MN_CMS
67            End If

68            TradesJuliaFormat = GetTradesInJuliaFormat(FilterBy1, _
                  Filter1Value, _
                  FilterBy2, _
                  Filter2Value, _
                  IncludeFutureTrades, _
                  PortfolioAgeing, _
                  FlipTrades, _
                  Numeraire, _
                  WithFxTrades, _
                  WithRatesTrades, _
                  TradesScaleFactor, _
                  CurrenciesToInclude, _
                  False, TC, twb, shFutureTrades, AnchorDate)

69        End If

          Dim OurValues

70        m_HeaderDiff = "Unlikely"

71        If WithValues Then
72            If sNRows(TradesJuliaFormat) = 1 Then
73                WithValues = False
74            End If
75        End If

76        If WithValues Then

77            OurValues = ThrowIfError(PortfolioValueHW(TradesJuliaFormat, ShockedModelName, "EUR", True, False))
78            OurValues = OneDToTwoD(OurValues)

79            If Choice = chExtraTrades Then

80                NumValueColumns = 1
81                ValueColumns = sArrayStack(m_HeaderHW, OurValues)

82                TradesJuliaFormat = sArrayRange(sSubArray(TradesJuliaFormat, 1, 1, , InsertValuesAfterCol), ValueColumns, sSubArray(TradesJuliaFormat, 1, InsertValuesAfterCol + 1))

83            Else
                  Dim CalypsoMultiplier As Double
                  Dim ChooseVector As Variant

84                ValuesCalypso = GetColumnFromTradesWorkbook("Pricer NPV", IncludeFutureTrades, PortfolioAgeing, WithFxTrades, _
                      WithRatesTrades, twb, shFutureTrades, AnchorDate)
85                ChooseVector = ChooseVectorFromFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value, _
                      IncludeFutureTrades, PortfolioAgeing, True, WithFxTrades, _
                      WithRatesTrades, CurrenciesToInclude, TC, twb, shFutureTrades, AnchorDate)
86                ValuesCalypso = sMChoose(ValuesCalypso, ChooseVector)
87                CalypsoMultiplier = TradesScaleFactor * IIf(FlipTrades, -1, 1)
88                If CalypsoMultiplier <> 1 Then ValuesCalypso = sArrayIf(sArrayIsNumber(ValuesCalypso), sArrayMultiply(ValuesCalypso, CalypsoMultiplier), ValuesCalypso)

89                EURNotionals = EURNotional(TradesJuliaFormat, ModelBareBones)

90                ValueDiff = sArrayIf(sArrayIsNumber(EURNotionals), sArrayMultiply(sArrayDivide(sArraySubtract(OurValues, ValuesCalypso), sArrayAbs(EURNotionals)), 10000), EURNotionals)
91                ValueDiff = sArrayIf(sArrayEquals(ValueDiff, "#Non-number found!"), Empty, ValueDiff)

92                If WithValues Then
93                    m_HeaderDiff = "HW-MTM bp"
94                    ValueHeaders = sArrayRange(m_HeaderHW, m_HeaderMTM, m_HeaderDiff)
95                    ValueColumns = sArrayStack(ValueHeaders, sArrayRange(OurValues, ValuesCalypso, ValueDiff))
96                End If
97                NumValueColumns = sNCols(ValueColumns)
98                TradesJuliaFormat = sArrayRange(sSubArray(TradesJuliaFormat, 1, 1, , InsertValuesAfterCol), ValueColumns, sSubArray(TradesJuliaFormat, 1, InsertValuesAfterCol + 1))
99            End If
100       End If

101       Set ws = shTradesViewer
102       ws.Visible = xlSheetVisible
103       Application.GoTo ws.Cells(1, 1)
104       Set SPH = CreateSheetProtectionHandler(ws)
105       ClearOutTradeValuesSheet
          Dim Parameters

106       If Choice = chExtraTrades Then

107           Parameters = sArrayStack(sArrayRange("Parameters:", Empty), _
                  sArrayRange("IncludeExtraTrades", True), _
                  sArrayRange("PortfolioAgeing", PortfolioAgeing), _
                  sArrayRange("FxShock", 1), _
                  sArrayRange("FxVolShock", 1), _
                  sArrayRange("TradesScaleFactor", 1))

108       Else

109           Parameters = sArrayStack(sArrayRange("Parameters:", Empty), _
                  sArrayRange("FilterBy1", FilterBy1), _
                  sArrayRange("Filter1Value", Filter1Value), _
                  sArrayRange("FilterBy2", FilterBy2), _
                  sArrayRange("IncludeExtraTrades", False), _
                  sArrayRange("PortfolioAgeing", PortfolioAgeing), _
                  sArrayRange("IncludeAssetClasses", IncludeAssetClasses), _
                  sArrayRange("CurrenciesToInclude", CurrenciesToInclude), _
                  sArrayRange("FxShock", FxShock), _
                  sArrayRange("FxVolShock", FxVolShock), _
                  sArrayRange("UseHistoricalFxVol", UseHistoricalFxVol), _
                  sArrayRange("TradesScaleFactor", TradesScaleFactor))

110           If FlipTrades Then
111               Parameters = sArrayStack(Parameters, sArrayRange("Trades are ""flipped"",i.e. mirror of trades in the trades workbook.", Empty))
112           End If

113           Parameters = CleanUpPromptArray(Parameters, False)       'remove uninteresting values
114       End If

115       Set Target = ws.Cells(3, 1).Resize(sNRows(Parameters), sNCols(Parameters))
116       With Target
117           .Value = sArrayExcelString(Parameters)
118           ws.Names.Add "Parameters", .offset(0)
119           AddGreyBorders .offset(0), True
120           .HorizontalAlignment = xlHAlignLeft
121       End With

122       Set Target = Target.offset(Target.Rows.Count + 4).Resize(sNRows(TradesJuliaFormat), sNCols(TradesJuliaFormat))

123       With Target
124           .Value = sArrayExcelString(TradesJuliaFormat)
125           AddGreyBorders .offset(0)
126           SetCommentsInHeaders .Rows(2), PortfolioAgeing, WithValues, FxShock, FxVolShock, TradesScaleFactor

127           If .Rows.Count > 1 Then
128               ws.Names.Add "TheData", .offset(1).Resize(.Rows.Count - 1)
129           End If

130           ws.Names.Add "TheHeaders", .Rows(1)
131           For Each c In .Rows(1).Cells
132               If InStr(c.Value, "Date") > 0 Then
133                   .Columns(c.Column - .Column + 1).NumberFormat = "dd-mmm-yyyy"
134               ElseIf InStr(c.Value, "Notional") > 0 Then
135                   .Columns(c.Column - .Column + 1).NumberFormat = "#,##0;[Red]-#,##0"
136               Else
137                   Select Case c.Value
                          Case m_HeaderHW, m_HeaderMTM, m_HeaderDiff
138                           .Columns(c.Column - .Column + 1).NumberFormat = "#,##0;[Red]-#,##0"
139                   End Select
140               End If
141           Next c
              Dim RangeToAutoFit As Range
142           Set RangeToAutoFit = ws.Range("Parameters")
143           Set RangeToAutoFit = RangeToAutoFit.Resize(RangeToAutoFit.Rows.Count - 1, 1)
144           Set RangeToAutoFit = Application.Union(.offset(0), RangeToAutoFit)
145           RangeToAutoFit.Columns.AutoFit

146           For Each c In .Rows(2).Cells
147               If c.ColumnWidth > 20 Then c.ColumnWidth = 20
148           Next c
149           With .Rows(-1)        'Add Filters
150               AddGreyBorders .offset(0)
151               CayleyFormatAsInput .offset(0)
152               .NumberFormat = "@"
153               .HorizontalAlignment = xlHAlignCenter
154               .Parent.Names.Add "TheFilters", .offset(0)
155           End With
156           AddSortButtons .Rows(0), 1
157           .Rows(1).Font.Bold = True
158           AddGroupingButtonToRange .Cells(-2, InsertValuesAfterCol + NumValueColumns).Resize(1, .Columns.Count - (InsertValuesAfterCol + NumValueColumns - 1)), True
159           With .Cells(-2, .Columns.Count + 1).Resize(3, 1)
160               .Value = sArrayStack(" <-click for more columns", " <-double-click to filter", " <-click to sort")
161               .Font.Color = g_Col_GreyText
162           End With

              'Set up the cell that displays the number of trades shown or filtered
163           With .Cells(-2, 1)
164               .Parent.Names.Add "NumTrades", .offset(0)
165               NumTrades = sNRows(TradesJuliaFormat) - 2
166               .Font.Color = g_Col_GreyText
167               If NumTrades = 0 Then
168                   .Value = "No trades to show"
169               ElseIf NumTrades = 1 Then
170                   .Value = "One trade shown"
171               Else
172                   .Value = "All " & Format(NumTrades, "#,###") & " trades shown."
173               End If
174           End With

175       End With

176       GroupingButtonDoAllOnSheet shTradesViewer, False
177       ActiveWindow.DisplayGridlines = False
178       ActiveWindow.DisplayHeadings = False

EarlyExit:
179       Set ws = Nothing
180       Set Mwb = Nothing
181       Set twb = Nothing
182       gBlockChangeEvent = oldBlockChange

          'Line below triggers change event that sets formatting and cell validation for TheFilters
183       If IsInCollection(shTradesViewer.Names, "TheFilters") Then
184           shTradesViewer.Range("TheFilters").ClearContents
185       End If

186       Exit Sub
ErrHandler:
187       SomethingWentWrong "#ShowTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
188       GoTo EarlyExit
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetCommentsInHeaders
' Author    : Philip Swannell
' Date      : 19-Oct-2016
' Purpose   : Set comments in the header row
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SetCommentsInHeaders(Headers As Range, PortfolioAgeing, DoHWPV As Boolean, FxShock, FxVolShock, TradesScaleFactor)

          Dim Comment As String
          Dim HeadersTP As Variant
          Dim MatchID As Variant

1         On Error GoTo ErrHandler
2         HeadersTP = sArrayTranspose(Headers.Value)

3         If PortfolioAgeing <> 0 Then
4             MatchID = sMatch("EndDate", HeadersTP)
5             If IsNumber(MatchID) Then
6                 Comment = "Because PortfolioAgeing is not zero the dates shown in this column are adjusted."
7                 SetCellComment Headers.Cells(1, MatchID), Comment, True
8             End If
9         End If

10        If DoHWPV Then
11            MatchID = sMatch(m_HeaderHW, HeadersTP)
12            If IsNumber(MatchID) Then
13                Comment = "PV in EUR from Banks's perspective. Model = Hull-White. Data from market data workbook."
14                If FxShock <> 1 Or FxVolShock <> 1 Then
15                    Comment = Comment & " Shocks applied to Fx rates and/or Fx vols."
16                End If
17                SetCellComment Headers.Cells(1, MatchID), Comment, True
18            End If
19        End If

20        If DoHWPV Then
21            MatchID = sMatch(m_HeaderDiff, HeadersTP)
22            If IsNumber(MatchID) Then
                  Dim relevantHeader As String
23                relevantHeader = m_HeaderHW
24                Comment = "'" & relevantHeader & "' minus '" & m_HeaderMTM & "' in basis points of notional."
25                If PortfolioAgeing <> 0 Then
26                    Comment = Comment & " NB: PortfolioAgeing has affected '" & relevantHeader & "' but has not affected '" & m_HeaderMTM & "'."
27                End If
28                SetCellComment Headers.Cells(1, MatchID), Comment, True
29            End If

30            MatchID = sMatch(m_HeaderMTM, HeadersTP)
31            If IsNumber(MatchID) Then
32                Comment = "PV in EUR from Banks's perspective, from 'MTM' column of the trades workbook"
33                If TradesScaleFactor <> 1 Then
34                    Comment = Comment & " multiplied by TradesScaleFactor (" & CStr(TradesScaleFactor) & ")."
35                Else
36                    Comment = Comment & "."
37                End If
38                SetCellComment Headers.Cells(1, MatchID), Comment, True
39            End If
40        End If

41        Exit Sub
ErrHandler:
42        Throw "#SetCommentsInHeaders (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

