Attribute VB_Name = "modFormatTrades"
'---------------------------------------------------------------------------------------
' Module    : modFormatTrades
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Methods for setting the cell formats for trades held on the Portfolio
'             sheet (or on other sheets)
'---------------------------------------------------------------------------------------
Option Explicit

Sub TestFormatTradesRange()
1         On Error GoTo ErrHandler
2         FormatTradesRange

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestFormatTradesRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormatTradesRange
' Author    : Philip Swannell
' Date      : 17-Nov-2015
' Purpose   : Applies cell formatting, cell validation and re-asserts formulas to the
'             trades in the Portfolio sheet (or potentially on other sheets)
'             Also morphs StartDate and EndDate from e.g. 10 to a date ten years from start date or AnchorDate
'             If SelectedCells is passed then only trades that intersect that selection are formatted
'---------------------------------------------------------------------------------------
Sub FormatTradesRange(Optional SelectedCells As Range, Optional TradesRange As Range)
          Dim CopyOfErr As String
          Dim CountRepeatsRet As Variant
          Dim EndDateCell As Range
          Dim ExSH As SolumAddin.clsExcelStateHandler
          Dim formulaShouldBe
          Dim i As Long
          Dim j As Long
          Dim NumCols As Long
          Dim NumTrades As Long
          Dim OldBCE As Boolean
          Dim SomeTradesOnly As Boolean
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim StartDateCell As Range
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim TargetCell As Range
          Dim ThisRange As Range
          Dim TradeType As String

1         On Error GoTo ErrHandler

2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         Set SUH = CreateScreenUpdateHandler()
5         Set ExSH = CreateExcelStateHandler(PreserveViewport:=True)

6         CalculatePortfolioSheet
7         shHiddenSheet.Calculate

8         If TradesRange Is Nothing Then
9             Set TradesRange = getTradesRange(NumTrades)
10            With TradesRange.Cells(TradesRange.Rows.Count + IIf(NumTrades = 0, 0, 1), 2)
11                If .Value <> gDoubleClickPrompt Then
12                    Set SPH = CreateSheetProtectionHandler(TradesRange.Parent)
13                    .Value = gDoubleClickPrompt
14                    .Font.Color = RangeFromSheet(shPortfolio, "NumTradesShown").Font.Color
15                    .HorizontalAlignment = xlHAlignCenter
16                End If
17            End With
18            If NumTrades = 0 Then GoTo EarlyExit
19        End If

20        If Not SelectedCells Is Nothing Then
21            SomeTradesOnly = True
22            Set TradesRange = Application.Intersect(TradesRange, SelectedCells.EntireRow)
23            If TradesRange Is Nothing Then GoTo EarlyExit
24        End If

25        If SPH Is Nothing Then Set SPH = CreateSheetProtectionHandler(TradesRange.Parent)
26        NumCols = TradesRange.Columns.Count

          Dim AnchorDate As Long
          Dim DoRoll As Boolean
          Dim EndRollM As Long
          Dim EndRollY As Long
          Dim StartRollM As Long
          Dim StartRollY As Long
          Dim VF As String

          'For convenience, allow user to enter integers (5 equivalent to "5Y") or strings such as "5Y" or "18M" into both start date and end date columns...
27        For Each StartDateCell In TradesRange.Columns(gCN_StartDate).Cells
28            Set EndDateCell = StartDateCell.Offset(0, gCN_EndDate - gCN_StartDate)
29            ParseLazyDate StartDateCell.Value2, True, StartRollY, StartRollM, DoRoll
30            If DoRoll Then
31                If AnchorDate = 0 Then AnchorDate = RangeFromMarketDataBook("Config", "AnchorDate").Value2
32                StartDateCell.Value = Application.WorksheetFunction.EDate(AnchorDate, 12 * StartRollY + StartRollM)
33            End If
34            ParseLazyDate EndDateCell.Value2, False, EndRollY, EndRollM, DoRoll
35            If DoRoll Then
36                If IsNumber(StartDateCell.Value2) Then
37                    EndDateCell.Value = Application.WorksheetFunction.EDate(StartDateCell.Value2, 12 * EndRollY + EndRollM)
38                Else
39                    VF = EndDateCell.Offset(0, gCN_TradeType - gCN_EndDate).Value
40                    If VF = "FxOption" Or VF = "FxForward" Then
41                        If AnchorDate = 0 Then AnchorDate = RangeFromMarketDataBook("Config", "AnchorDate").Value2
42                        EndDateCell.Value = Application.WorksheetFunction.EDate(AnchorDate, 12 * EndRollY + EndRollM)
43                    End If
44                End If
45            End If
46        Next

          'Allow "k", "m" and "b" as abbreviations for thousand, million and billion
          Dim Cell As Range
          Dim R As New RegExp
          Dim tmp As Variant
47        With R
48            .IgnoreCase = True
49            .Pattern = "^[0-9\.\-kmb;]*$"
50            .Global = False
51        End With

52        For Each Cell In Application.Union(TradesRange.Columns(gCN_Notional1), TradesRange.Columns(gCN_Notional2)).Cells
53            If VarType(Cell.Value) = vbString Then
54                If Not Cell.HasFormula Then
55                    If R.Test(Cell.Value) Then
56                        If (InStr(1, Cell.Value, "m", vbTextCompare) + InStr(1, Cell.Value, "b", vbTextCompare) + InStr(1, Cell.Value, "k", vbTextCompare)) > 0 Then
57                            tmp = UnabbreviateNotional(Cell.Value)
58                            If tmp <> Cell.Value Then
59                                Cell.Value = tmp
60                            End If
61                        End If
62                    End If
63                End If
64            End If
65        Next Cell

66        With TradesRange
67            .Validation.Delete
68            .Resize(, .Columns.Count - 1).ClearFormats
69            .Resize(, .Columns.Count - 1).HorizontalAlignment = xlHAlignCenter    'horizontal alignment of right-most column is dealt with by method UpdatePortfolioSheet
70            .VerticalAlignment = xlVAlignBottom
71            .Font.Color = Colour_BlueText
72            .Columns(gCN_TradeType).Font.ColorIndex = xlColorIndexAutomatic
73            .Columns(gCN_StartDate).NumberFormat = NF_Date
74            .Columns(gCN_EndDate).NumberFormat = NF_Date
75            .Columns(gCN_Notional1).NumberFormat = NF_Comma0dp
76            .Columns(gCN_Notional2).NumberFormat = NF_Comma0dp
77            .Columns(gCN_Rate1).NumberFormat = "0.000%;[Red]-0.000%"
78            .Columns(gCN_Rate2).NumberFormat = "0.000%;[Red]-0.000%"
79            .Columns(.Columns.Count - 2).NumberFormat = NF_Fx    'show 4 decimal places, or if >10 show 3 dp, or if >=100 show 2dp - we could use conditional fomatting for more control...
80            .Columns(.Columns.Count - 1).Resize(, 2).NumberFormat = NF_Comma0dp
81            .Columns(.Columns.Count - 2).Resize(, 3).Font.ColorIndex = xlColorIndexAutomatic
82            .Resize(, .Columns.Count - 3).Locked = False
83            .Columns(gCN_TradeType).Locked = True
84            If gDoValidation Then
85                SetValidationForCurrency .Columns(gCN_Ccy2)
86                SetValidationForCurrencyOrInflation .Columns(gCN_Ccy1), .Columns(gCN_TradeType)
87                SetValidationForFreq Application.Union(.Columns(gCN_Freq1), .Columns(gCN_Freq2))
88                SetValidationForDCT Application.Union(.Columns(gCN_DCT1), .Columns(gCN_DCT2))
89                SetValidationForBDC Application.Union(.Columns(gCN_BDC1), .Columns(gCN_BDC2))
90            End If
91            ConditionalFormattingForCounterparty .Columns(gCN_Counterparty)
92            AddGreyBorders .Offset(0)
93        End With

94        CountRepeatsRet = sCountRepeats(TradesRange.Columns(gCN_TradeType).Value, "CFH")

95        For i = 1 To sNRows(CountRepeatsRet)
96            TradeType = CountRepeatsRet(i, 1)
97            Set ThisRange = TradesRange.Rows(CountRepeatsRet(i, 2)).Resize(CountRepeatsRet(i, 3))
98            Select Case TradeType
                  Case "InterestRateSwap"
99                    For j = 1 To 2
100                       FormatLocked ThisRange.Columns(Choose(j, gCN_Ccy2, gCN_Notional2))
101                       If gDoValidation Then SetValidationForNotional ThisRange.Columns(gCN_Notional1)
102                   Next j
103                   If gDoValidation Then
104                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
105                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
106                   End If

107                   ThisRange.Columns(gCN_Ccy2).FormulaR1C1 = "=RC[" + CStr(gCN_Ccy1 - gCN_Ccy2) + "]"
108               Case "FxForward", "FxOption", "FxOptionStrip", "FxForwardStrip"
109                   For j = 1 To IIf(Left(TradeType, 9) = "FxForward", 11, 10)
110                       FormatNA ThisRange.Columns(Choose(j, gCN_StartDate, gCN_Rate1, gCN_Freq1, gCN_Rate2, gCN_LegType2, gCN_Freq2, gCN_DCT1, gCN_DCT2, gCN_BDC1, gCN_BDC2, gCN_LegType1))
111                   Next j
112                   If Right(TradeType, 5) = "Strip" Then
113                       ConditionalFormattingDoubleClickMe Application.Union(ThisRange.Columns(gCN_EndDate), ThisRange.Columns(gCN_Notional1), ThisRange.Columns(gCN_Notional2))
114                   Else
115                       If gDoValidation Then SetValidationForStartDate ThisRange.Columns(gCN_EndDate)
116                   End If
117               Case "Swaption"
118                   For j = 1 To 2
119                       FormatNA ThisRange.Columns(Choose(j, gCN_Ccy2, gCN_Notional2))
120                       ThisRange.Columns(gCN_Ccy2).FormulaR1C1 = "=RC[" + CStr(gCN_Ccy1 - gCN_Ccy2) + "]"
121                       ThisRange.Columns(gCN_Notional2).FormulaR1C1 = "=RC[" + CStr(gCN_Notional1 - gCN_Notional2) + "]"
122                   Next j
123                   For j = 1 To 2
124                       FormatLocked ThisRange.Columns(Choose(j, gCN_Rate2, gCN_LegType2))
125                   Next j
126                   If gDoValidation Then
127                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
128                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
129                   End If
130               Case "CrossCurrencySwap"
131                   If gDoValidation Then
132                       SetValidationForNotional ThisRange.Columns(gCN_Notional1)
133                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
134                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
135                   End If
136               Case "CapFloor"
137                   For j = 1 To 7
138                       FormatNA ThisRange.Columns(Choose(j, gCN_Ccy2, gCN_Notional2, gCN_Rate2, gCN_LegType2, gCN_Freq2, gCN_DCT2, gCN_BDC2))
139                   Next j
140                   If gDoValidation Then
141                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
142                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
143                   End If
144               Case "FixedCashflows"
145                   For j = 1 To 13
146                       FormatNA ThisRange.Columns(Choose(j, gCN_StartDate, gCN_Rate1, gCN_LegType1, gCN_Freq1, gCN_Ccy2, gCN_Notional2, gCN_Rate2, gCN_LegType2, gCN_Freq2, gCN_DCT1, gCN_DCT2, gCN_BDC1, gCN_BDC2))
147                   Next j
148                   ConditionalFormattingDoubleClickMe Application.Union(ThisRange.Columns(gCN_EndDate), ThisRange.Columns(gCN_Notional1))
149               Case "InflationZCSwap"
150                   For j = 1 To 5
151                       FormatNA ThisRange.Columns(Choose(j, gCN_DCT1, gCN_Ccy2, gCN_DCT2, gCN_Freq2, gCN_Freq1))
152                   Next j
153                   FormatLocked ThisRange.Columns(gCN_LegType2), "=IF(RC[" + CStr(gCN_LegType1 - gCN_LegType2) + "]=""Fixed"",""Index"",""Fixed"")"
154                   FormatLocked ThisRange.Columns(gCN_Notional2), "=RC[" + CStr(gCN_Notional1 - gCN_Notional2) + "]"
155                   FormatLocked ThisRange.Columns(gCN_BDC2), "=RC[" + CStr(gCN_BDC1 - gCN_BDC2) + "]"
156                   If gDoValidation Then
157                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
158                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
159                   End If
160               Case "InflationYoYSwap"
161                   FormatLocked ThisRange.Columns(gCN_Notional2), "=RC[" + CStr(gCN_Notional1 - gCN_Notional2) + "]"
162                   If gDoValidation Then
163                       SetValidationForStartDate ThisRange.Columns(gCN_StartDate)
164                       SetValidationForEndDate ThisRange.Columns(gCN_EndDate)
165                   End If
166           End Select
167       Next i

          'The code below loops through all the trades and applies formatting which may differ between two _
           examples of a given type of trade - necessary since the loop above handles "blocks" of trades of the same type
168       For i = 1 To TradesRange.Rows.Count
169           Select Case TradesRange.Cells(i, gCN_TradeType).Value
                  Case "InterestRateSwap"
170                   Set TargetCell = TradesRange.Cells(i, gCN_Notional2)
171                   If VarType(TradesRange.Cells(i, gCN_Notional1)) = vbString Then    ' This is an amortising trade
172                       With TargetCell
173                           If .Interior.ColorIndex <> xlColorIndexNone Then .Interior.ColorIndex = xlColorIndexNone
174                           If .Font.Color <> Colour_BlueText Then .Font.Color = Colour_BlueText
175                           If .Locked Then .Locked = False
176                           .HorizontalAlignment = xlHAlignLeft
177                           TradesRange.Cells(i, gCN_Notional1).HorizontalAlignment = xlHAlignLeft
178                       End With
179                   Else
                          'Not an amortising trade
180                       formulaShouldBe = "=" + Replace(TradesRange.Cells(i, gCN_Notional1).Address, "$", "")
181                       With TargetCell
182                           If .Formula <> formulaShouldBe Then
183                               .Formula = formulaShouldBe
184                           End If
185                           If .Font.Color <> Colour_GreyText Then .Font.Color = Colour_GreyText
186                           If .Interior.Color <> Colour_LightGrey Then .Interior.Color = Colour_LightGrey
187                           If Not .Locked Then .Locked = True
188                           .HorizontalAlignment = xlHAlignCenter
189                           TradesRange.Cells(i, gCN_Notional1).HorizontalAlignment = xlHAlignCenter
190                       End With
191                   End If
192               Case "CrossCurrencySwap"    ' only thing to do is set horizontal alignment of notionals
193                   For j = 1 To 2
194                       With TradesRange.Cells(i, Choose(j, gCN_Notional1, gCN_Notional2))
195                           If VarType(.Value) = vbString Then    'amortising trade
196                               .HorizontalAlignment = xlHAlignLeft
197                           Else    ' not amortising trade
198                               If .HorizontalAlignment <> xlHAlignCenter Then .HorizontalAlignment = xlHAlignCenter
199                           End If
200                       End With
201                   Next j
202               Case "InflationZCSwap"
203                   If LCase(TradesRange.Cells(i, gCN_LegType1)) = "fixed" Then
                          Dim GreyCol As Long
                          Dim WhiteCol As Long
204                       WhiteCol = gCN_Rate1: GreyCol = gCN_Rate2
205                   Else
206                       WhiteCol = gCN_Rate2: GreyCol = gCN_Rate1
207                   End If
208                   FormatUserEntered TradesRange.Cells(i, WhiteCol), True
209                   FormatIllegible TradesRange.Cells(i, GreyCol), True
210               Case "InflationYoYSwap"
211                   For j = 1 To 2
212                       Select Case TradesRange.Cells(i, Choose(j, gCN_LegType1, gCN_LegType2)).Value
                              Case "Index"
213                               FormatIllegible TradesRange.Cells(i, Choose(j, gCN_Rate1, gCN_Rate2)), True
                                  Dim BaseCCy As String
214                               BaseCCy = InflationIndexInfo(TradesRange.Cells(i, Choose(j, gCN_Ccy1, gCN_Ccy2)).Value, "BaseCurrency")
215                               If Not sIsErrorString(BaseCCy) Then
216                                   With TradesRange.Cells(i, Choose(j, gCN_Ccy2, gCN_Ccy1))
217                                       If .Value <> BaseCCy Then .Value = BaseCCy
218                                   End With
219                               End If
220                           Case Else
221                               FormatUserEntered TradesRange.Cells(i, Choose(j, gCN_Rate1, gCN_Rate2)), True
222                       End Select
223                   Next j
224           End Select
225       Next i

226       If Not SomeTradesOnly Then
227           If TradesRange.Parent Is shPortfolio Then
228               With RangeFromSheet(shPortfolio, "TheFilters")
229                   .Locked = False
230                   .HorizontalAlignment = xlHAlignCenter
231                   .NumberFormat = "@"
232                   AddGreyBorders .Offset(0)
233                   .WrapText = False
234               End With
235           End If
236       End If

237       If TradesRange.Parent Is shPortfolio Then
238           CalculatePortfolioSheet
239           SetTradesRangeColumnWidths
240       End If

EarlyExit:
241       gBlockChangeEvent = OldBCE
242       Exit Sub
ErrHandler:
243       CopyOfErr = "#FormatTradesRange (line " & CStr(Erl) + "): " & Err.Description & "!"
244       gBlockChangeEvent = OldBCE
245       Throw CopyOfErr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseLazyDate
' Author     : Philip Swannell
' Date       : 17-Jan-2018
' Purpose    : Parse a "lazily entered date" in terms of the number of months and years that the user intends to roll
'              from either the start date of the anchor date.
' Parameters :
'  Data    : Value2 property of the data the user has entered
'  AllowNeg: whether we allow negative rolls
'  RollY   : by reference, returns the number of years to roll
'  RollM   : by reference, returns the number of months to roll
'  DoRoll  : By reference, returns whether the date was entered in "lazy format" i.e. if we need to convert what the user
'            entered into a valid date by "rolling"
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseLazyDate(Data, AllowNeg As Boolean, ByRef RollY As Long, ByRef RollM As Long, ByRef DoRoll As Boolean)
          Dim LB As Long
          Const UB = 100
1         LB = IIf(AllowNeg, -100, 0)
2         RollM = 0: RollY = 0: DoRoll = False
3         If IsNumber(Data) Then
4             If LB <= Data And UB >= Data Then
5                 DoRoll = True
6                 RollY = Data
7             End If
8         ElseIf VarType(Data) = vbString Then
9             If Right(UCase(Data), 1) = "Y" Then
10                If IsNumeric(Left(Data, Len(Data) - 1)) Then
11                    RollY = CLng(Left(Data, Len(Data) - 1))
12                    If RollY >= LB And RollY <= UB Then DoRoll = True
13                End If
14            ElseIf Right(UCase(Data), 1) = "M" Then
15                If IsNumeric(Left(Data, Len(Data) - 1)) Then
16                    RollM = CLng(Left(Data, Len(Data) - 1))
17                    If RollM >= (LB * 12) And RollM <= (UB * 12) Then DoRoll = True
18                End If
19            End If
20        End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnabbreviateNotional
' Author    : Philip
' Date      : 27-Sep-2017
' Purpose   : UnAbbreviate a notional
'         eg. "10m" > 10000000
'       "1m;1m;1m"  > "1000000;1000000;1000000"
'           "2.25b" > 2250000000
'---------------------------------------------------------------------------------------
Private Function UnabbreviateNotional(NotionalString As String)
1         On Error GoTo ErrHandler
2         If InStr(NotionalString, ";") = 0 Then
3             UnabbreviateNotional = UnAbCore(NotionalString)
4         Else
              Dim i As Long
              Dim TmpArray() As String
              Dim TmpArray2() As String
5             TmpArray = VBA.Split(NotionalString, ";")
6             TmpArray2 = TmpArray
7             For i = LBound(TmpArray) To UBound(TmpArray)
8                 TmpArray2(i) = CStr(UnAbCore(TmpArray(i)))
9             Next i
10            UnabbreviateNotional = VBA.Join(TmpArray2, ";")
11        End If
12        Exit Function
ErrHandler:
13        Throw "#UnabbreviateNotional (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function UnAbCore(NotionalString As String)
          Dim tmp As Variant
1         On Error GoTo ErrHandler
2         tmp = Application.Evaluate("=" & Replace(Replace(Replace(LCase(NotionalString), "k", "*1000"), "m", "*1000000"), "b", "*1000000000"))
3         If IsNumber(tmp) Then
4             UnAbCore = tmp
5         Else
6             UnAbCore = NotionalString
7         End If
8         Exit Function
ErrHandler:
9         Throw "#UnAbCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub FormatIllegible(R As Range, TestFirst As Boolean)
1         On Error GoTo ErrHandler
2         With R
3             If TestFirst Then
4                 If .Interior.Color <> Colour_LightGrey Then .Interior.Color = Colour_LightGrey
5                 If .Font.Color <> Colour_LightGrey Then .Font.Color = Colour_LightGrey
6                 If Not .Locked Then .Locked = True
7             Else
8                 .Interior.Color = Colour_LightGrey
9                 .Font.Color = Colour_LightGrey
10                .Locked = True
11            End If
12        End With
13        Exit Sub
ErrHandler:
14        Throw "#FormatIllegible (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub FormatUserEntered(R As Range, TestFirst As Boolean)
1         On Error GoTo ErrHandler
2         With R
3             If TestFirst Then
4                 If .Interior.Color <> 16777215 Then .Interior.Color = 16777215
5                 If .Font.Color <> Colour_BlueText Then .Font.Color = Colour_BlueText
6                 If .Locked Then .Locked = False
7             Else
8                 .Interior.Color = 16777215
9                 .Font.Color = Colour_BlueText
10                .Locked = False
11            End If
12        End With
13        Exit Sub
ErrHandler:
14        Throw "#FormatUserEntered (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub FormatLocked(R As Range, Optional FormulaR1C1 As String)
1         On Error GoTo ErrHandler
2         If FormulaR1C1 <> "" Then R.FormulaR1C1 = FormulaR1C1
3         R.Font.Color = Colour_GreyText
4         R.Interior.Color = Colour_LightGrey
5         R.Locked = True
6         Exit Sub
ErrHandler:
7         Throw "#FormatLocked (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub FormatNA(R As Range)
1         On Error GoTo ErrHandler
2         R.Font.Color = Colour_LightGreyText
3         R.Interior.Color = Colour_LightGrey
4         R.Value = "N/A"
5         R.Locked = True
6         Exit Sub
ErrHandler:
7         Throw "#FormatNA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FormatTradeTemplates()
1         On Error GoTo ErrHandler
2         FormatTradesRange , sExpandDown(RangeFromSheet(shHiddenSheet, "Headers").Rows(3))
3         Exit Sub
ErrHandler:
4         Throw "#FormatTradeTemplates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FilterTradesRange
' Author    : Philip Swannell
' Date      : 26-Apr-2016
' Purpose   : Display only some of the trades according to filters entered by the user
'---------------------------------------------------------------------------------------
Sub FilterTradesRange()
1         On Error GoTo ErrHandler
          Dim Message As String
          Dim SPH As SolumAddin.clsSheetProtectionHandler

2         FilterRangeByHidingRows RangeFromSheet(shPortfolio, "TheFilters"), getTradesRange(0), "trade", Message, _
              , sArrayTranspose(RangeFromSheet(shHiddenSheet, "RegKeys"))

3         If RangeFromSheet(shPortfolio, "NumTradesShown").Value <> Message Then
4             Set SPH = CreateSheetProtectionHandler(shPortfolio)
5             RangeFromSheet(shPortfolio, "NumTradesShown").Value = Message
6         End If

7         Exit Sub
ErrHandler:
8         Throw "#FilterTradesRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function SetValidationForNotional(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
5             .InputMessage = "Double-click for amortising" + vbLf + _
                  "K  = 000" + vbLf + _
                  "M = 000,000" + vbLf '+ _
                  "B  = 000,000,000" + vbLf ' B works as billion but mentioning that in the validation message is a bit verbose
6         End With
7         Exit Function
ErrHandler:
8         Throw "#SetValidationForNotional (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetValidationForStartDate(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
5             .InputMessage = "Date or text such as '5Y', '18M'"
6         End With
7         Exit Function
ErrHandler:
8         Throw "#SetValidationForStartDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetValidationForEndDate(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
5             .InputMessage = "Date or text such as '5Y', '18M'"
6         End With
7         Exit Function
ErrHandler:
8         Throw "#SetValidationForEndDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetValidationForCurrencyOrInflation
' Author    : Philip Swannell
' Date      : 21-Apr-2017
' Purpose   : Set validation on a column (such as the "Ccy 1" column) that contains some
'             cells that should contain currencies and other cells that should contain
'             inflation indices
'---------------------------------------------------------------------------------------
Private Function SetValidationForCurrencyOrInflation(R As Range, TradeTypeRange As Range)
          Dim CountRepeatsRet As Variant
          Dim i As Long
          Dim NR As Long
          Dim ThisChunk As Range
          Dim TradeTypes As Variant
          Dim ValTypes As Variant
1         On Error GoTo ErrHandler

2         TradeTypes = TradeTypeRange.Value
3         Force2DArray TradeTypes
4         NR = sNRows(TradeTypes)
5         ValTypes = sReshape("", NR, 1)

6         For i = 1 To NR
7             Select Case TradeTypes(i, 1)
                  Case "InflationZCSwap"
8                     ValTypes(i, 1) = "I"
9                 Case "InflationYoYSwap"
10                    ValTypes(i, 1) = "CorI"
11                Case Else
12                    ValTypes(i, 1) = "C"
13            End Select
14        Next i

15        CountRepeatsRet = sCountRepeats(ValTypes, "CFH")
16        For i = 1 To sNRows(CountRepeatsRet)
17            Set ThisChunk = R.Cells(CountRepeatsRet(i, 2), 1).Resize(CountRepeatsRet(i, 3))
18            Select Case CountRepeatsRet(i, 1)
                  Case "I"
19                    SetValidationForInflation ThisChunk
20                Case "C"
21                    SetValidationForCurrency ThisChunk
22                Case "CorI"
23                    SetValidationCurrencyOrInflation ThisChunk
24            End Select
25        Next i

26        Exit Function
ErrHandler:
27        Throw "#SetValidationForCurrencyOrInflation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetValidationForInflation(R As Range)
1         On Error GoTo ErrHandler
          Static InflationList As String

2         If InflationList = "" Then
3             InflationList = sConcatenateStrings(SupportedInflationIndices())
4         End If

5         With R.Validation
6             .Delete
7             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=InflationList
8             .IgnoreBlank = True
9             .InCellDropdown = False
10            .InputTitle = ""
11            .ErrorTitle = gProjectName
12            .InputMessage = ""
13            .ErrorMessage = _
                  "That's not a valid inflation index." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
14            .ShowInput = True
15            .ShowError = True
16        End With

17        Exit Function
ErrHandler:
18        Throw "#SetValidationForInflation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetValidationForCurrency(R As Range)
1         On Error GoTo ErrHandler
          Static CurrencyList As String

2         If CurrencyList = "" Then
3             CurrencyList = sConcatenateStrings(sCurrencies(False, True))
4         End If

5         With R.Validation
6             .Delete
7             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=CurrencyList
8             .IgnoreBlank = True
9             .InCellDropdown = False
10            .InputTitle = ""
11            .ErrorTitle = gProjectName
12            .InputMessage = ""
13            .ErrorMessage = _
                  "That's not a valid Currency." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
14            .ShowInput = True
15            .ShowError = True
16        End With

17        Exit Function
ErrHandler:
18        Throw "#SetValidationForCurrency (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SetValidationCurrencyOrInflation(R As Range)
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
          Static Allowed As String

3         If Allowed = "" Then
4             Allowed = sConcatenateStrings(sSortedArray(sArrayStack(sCurrencies(False, True), SupportedInflationIndices())))
5         End If

6         With R.Validation
7             .Delete
8             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=Allowed
9             .IgnoreBlank = True
10            .InCellDropdown = False
11            .InputTitle = ""
12            .ErrorTitle = gProjectName
13            .InputMessage = ""
14            .ErrorMessage = _
                  "That's not a valid Currency or Inflation Index." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
15            .ShowInput = True
16            .ShowError = True
17        End With
18        Exit Function
ErrHandler:
19        Throw "#SetValidationCurrencyOrInflation (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

'---------------------------------------------------------------------------------------
' Procedure : SetValidationForDCT
' Author    : Philip Swannell
' Date      : 10-May-2016
' Purpose   : Add data validation to cells that contain day count types
'---------------------------------------------------------------------------------------
Private Function SetValidationForDCT(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=sConcatenateStrings(sSupportedDCTs())
5             .IgnoreBlank = True
6             .InCellDropdown = False
7             .InputTitle = ""
8             .ErrorTitle = gProjectName
9             .InputMessage = ""
10            .ErrorMessage = _
                  "That's not a valid day count type." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
11            .ShowInput = True
12            .ShowError = True
13        End With

14        Exit Function
ErrHandler:
15        Throw "#SetValidationForDCT (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetValidationForBDC
' Author    : Philip Swannell
' Date      : 10-May-2016
' Purpose   : Add data validation to cells that contain business day conventions
'---------------------------------------------------------------------------------------
Private Function SetValidationForBDC(R As Range)
1         On Error GoTo ErrHandler
          Static BDCs As String
2         If BDCs = "" Then BDCs = sConcatenateStrings(SupportedBDCs())

3         With R.Validation
4             .Delete
5             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:=BDCs
6             .IgnoreBlank = True
7             .InCellDropdown = False
8             .InputTitle = ""
9             .ErrorTitle = gProjectName
10            .InputMessage = ""
11            .ErrorMessage = _
                  "That's not a valid business day convention." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
12            .ShowInput = True
13            .ShowError = True
14        End With

15        Exit Function
ErrHandler:
16        Throw "#SetValidationForBDC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetValidationForFreq
' Author    : Philip Swannell
' Date      : 27-Apr-2016
' Purpose   : Adding data validation to cells that contain payment frequency.
'            CHANGING THIS FUNCTION? then also make equivalent change to method sParseFrequencyString
'---------------------------------------------------------------------------------------
Private Function SetValidationForFreq(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                  Operator:=xlBetween, Formula1:="Annual,Ann,A,Semi annual,Semi-annual,Semi,S,Quarterly,Quarter,Quart,Q,Monthly,Month,M"
5             .IgnoreBlank = True
6             .InCellDropdown = False
7             .InputTitle = ""
8             .ErrorTitle = gProjectName
9             .InputMessage = ""
10            .ErrorMessage = _
                  "That's not a valid payment frequency." & Chr(10) & "" & Chr(10) & "Double-click to see what is allowed."
11            .ShowInput = True
12            .ShowError = True
13        End With

14        Exit Function
ErrHandler:
15        Throw "#SetValidationForFreq (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ConditionalFormattingForCounterparty(R As Range)
1         On Error GoTo ErrHandler
2         With R
3             .FormatConditions.Delete
4             .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                  Formula1:="=""" & gWHATIF & """"
5             .FormatConditions(.FormatConditions.Count).SetFirstPriority
6             With .FormatConditions(1).Font
7                 .Color = RGB(255, 0, 0)
8                 .TintAndShade = 0
9             End With
10            .FormatConditions(1).StopIfTrue = False
11            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                  Formula1:="=""??"""
12            .FormatConditions(.FormatConditions.Count).SetFirstPriority
13            With .FormatConditions(1).Interior
14                .PatternColorIndex = xlAutomatic
15                .Color = Colour_LightYellow
16                .TintAndShade = 0
17            End With
18            .FormatConditions(1).StopIfTrue = False
19        End With

20        Exit Function
ErrHandler:
21        Throw "#ConditionalFormattingForCounterparty (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ConditionalFormattingDoubleClickMe(R As Range)
1         On Error GoTo ErrHandler
2         With R
3             .FormatConditions.Delete
4             .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                  Formula1:="=""Double-click me!"""
5             .FormatConditions(.FormatConditions.Count).SetFirstPriority
6             With .FormatConditions(1).Interior
7                 .PatternColorIndex = xlAutomatic
8                 .Color = Colour_LightYellow
9                 .TintAndShade = 0
10            End With
11            .FormatConditions(1).StopIfTrue = False
12        End With

13        Exit Function
ErrHandler:
14        Throw "#ConditionalFormattingDoubleClickMe (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


