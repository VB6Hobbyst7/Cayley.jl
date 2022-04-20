Attribute VB_Name = "modDoubleclick"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : TradeDoubleClickHandler
' Author    : Philip Swannell
' Date      : 23-Nov-2015
' Purpose   : Handles double-click event on trades on the Portfolio sheet.
'---------------------------------------------------------------------------------------
Sub TradeDoubleClickHandler(Target As Range, TradeRange As Range, TradeType As String, TradeAttribute As String, ByRef Cancel As Boolean)
          Dim Alternatives As Variant
          Dim CurrentChoice
          Dim EnlargedTarget As Range
          Dim HelpMethodName As String
          Dim IsHyphenated As Boolean
          Dim originalValue As Variant
          Dim Res As Variant
          Dim shCDS As Worksheet
          Dim TopText As String
          Dim V As Variant

1         On Error GoTo ErrHandler

2         If Selection.Cells.Count > 1 Then    'can only happen when user triggers via Alt-Q (Alt X on Office 2016)
3             Set EnlargedTarget = Application.Intersect(Selection, getTradesRange(0))
4             If Application.Intersect(EnlargedTarget.EntireColumn, shPortfolio.Rows(1)).Cells.Count > 1 Then
5                 Throw "You cannot edit more than one column at a time.", True
6             End If
7             Set EnlargedTarget = UnhiddenRowsInRange(EnlargedTarget)
8             If Not sEquals(False, EnlargedTarget.Locked) Then
9                 Throw "Some of those cells cannot be edited.", True
10            End If
11        Else
12            Set EnlargedTarget = Target
13        End If

14        TopText = "Select " + TradeAttribute

15        Set shCDS = OpenMarketWorkbook(True, False).Worksheets("Credit")

16        CurrentChoice = Target.Cells(1, 1).Value

17        Select Case LCase(TradeAttribute)
              Case "start date"
18                If Right(TradeType, 5) <> "Strip" And TradeType <> "FixedCashflows" Then
19                    DatePicker Alternatives, Year(Date) - 5, Year(Date) + 5, Target, True, True
20                    If sIsErrorString(Alternatives) Then Alternatives = Empty
21                End If
22            Case "end date"
23                If Right(TradeType, 5) <> "Strip" And TradeType <> "FixedCashflows" Then
24                    DatePicker Alternatives, Year(Date), Year(Date) + 20, Target, True, True
25                    If sIsErrorString(Alternatives) Then Alternatives = Empty
26                End If
27            Case "cpty", "counterparty"

28                Alternatives = sSortedArray(CounterpartiesFromMarketBook(True))

29            Case "ccy 1"
                  'Note that Alternatives may be overwritten by later code in this method, e.g. for inflation trades
30                Alternatives = CurrenciesSupported(True, True)
31                IsHyphenated = True
32        End Select
33        Select Case LCase(TradeType + "|" + TradeAttribute)
              Case "inflationzcswap|ccy 1"
34                TopText = "Select Inflation Index"
35                Alternatives = SupportedInflationIndices()
36                IsHyphenated = False
37            Case "inflationyoyswap|ccy 1", "inflationyoyswap|ccy 2"
38                Select Case CStr(Target.Offset(0, gCN_LegType1 - gCN_Ccy1).Value)
                      Case "Index"
39                        TopText = "Select Inflation Index"
40                        Alternatives = SupportedInflationIndices()
41                        IsHyphenated = False
42                    Case Else
                          Dim BaseCCy
43                        BaseCCy = InflationIndexInfo(TradeRange.Cells(1, IIf(TradeAttribute = "Ccy 1", gCN_Ccy2, gCN_Ccy1)), "BaseCurrency")
44                        If sIsErrorString(BaseCCy) Then
45                            Alternatives = CurrenciesSupported(True, True)
46                            IsHyphenated = True
47                        Else
48                            Alternatives = BaseCCy
49                            IsHyphenated = False
50                        End If
51                End Select
52            Case "interestrateswap|leg type 1", "interestrateswap|leg type 2", "crosscurrencyswap|leg type 2", "crosscurrencyswap|leg type 2"
53                Alternatives = SupportedIRLegTypes()
54            Case "inflationzcswap|leg type 1"
55                Alternatives = sArrayStack("Index", "Fixed")
56            Case "inflationyoyswap|leg type 1", "inflationyoyswap|leg type 2"
57                Alternatives = sArrayStack("Index", "Fixed", "Floating")
58            Case "interestrateswap|freq 1", "interestrateswap|freq 2", "crosscurrencyswap|freq 1", "crosscurrencyswap|freq 2", "swaption|freq 1", "swaption|freq 2", "capfloor|freq 1", "inflationyoyswap|freq 1", "inflationyoyswap|freq 2"
59                IsHyphenated = False
60                Alternatives = sArrayStack("Annual", "Semi annual", "Quarterly", "Monthly")
61                CurrentChoice = sParseFrequencyString(CStr(CurrentChoice), False, False)
62            Case "interestrateswap|dct 1", "interestrateswap|dct 2", "crosscurrencyswap|dct 1", "crosscurrencyswap|dct 2", "swaption|dct 1", "swaption|dct 2", "capfloor|dct 1"
63                If Target.Offset(0, gCN_LegType1 - gCN_DCT1).Value = "Floating" Or TradeType = "CapFloor" Then
64                    Alternatives = sArrayStack("A/360", "A/365F")
65                Else
66                    Alternatives = sSupportedDCTs()
67                    HelpMethodName = "'" + ThisWorkbook.Name + "'!ShowHelpOnDayCounts"
68                End If
69            Case "inflationyoyswap|dct 1", "inflationyoyswap|dct 2"
70                Select Case Target.Offset(0, gCN_LegType1 - gCN_DCT1).Value
                      Case "Fixed"
71                        Alternatives = sSupportedDCTs()
72                        HelpMethodName = "'" + ThisWorkbook.Name + "'!ShowHelpOnDayCounts"
73                    Case "Index"
74                        Alternatives = "ActB/ActB"
75                    Case "Floating"
76                        Alternatives = sArrayStack("A/360", "A/365F")
77                    Case Else
78                        Alternatives = sSupportedDCTs()
79                        HelpMethodName = "'" + ThisWorkbook.Name + "'!ShowHelpOnDayCounts"
80                End Select
81            Case "interestrateswap|bdc 1", "interestrateswap|bdc 2", "crosscurrencyswap|bdc 1", "crosscurrencyswap|bdc 2", "swaption|bdc 1", "swaption|bdc 2", "capfloor|bdc 1", "inflationyoyswap|bdc 1", "inflationyoyswap|bdc 2", "inflationzcswap|bdc 1"
82                Alternatives = SupportedBDCs()
83            Case "fxoption|leg type 1", "fxoptionstrip|leg type 1"
84                IsHyphenated = True
                  Dim ThisCcy1 As String
85                ThisCcy1 = Target.Offset(, gCN_Ccy1 - gCN_LegType1).Value
86                If Len(ThisCcy1) <> 3 Then ThisCcy1 = "Ccy1"
87                TopText = "Select option style"
88                Alternatives = sArrayStack("BuyCall - Buy Call on " + ThisCcy1, _
                      "SellCall - Sell Call on " + ThisCcy1, _
                      "BuyPut - Buy Put on " + ThisCcy1, _
                      "SellPut - Sell Put on " + ThisCcy1)
89            Case "swaption|leg type 1"
90                Alternatives = sArrayStack("BuyReceivers", "SellReceivers", "BuyPayers", "SellPayers")
91                TopText = "Select option style"
92            Case "capfloor|leg type 1"
93                TopText = "Select option style"
94                Alternatives = sArrayStack("BuyCap", "SellCap", "BuyFloor", "SellFloor")
95            Case "crosscurrencyswap|ccy 2", "fxforward|ccy 2", "fxoption|ccy 2", "fxoptionstrip|ccy 2"
96                Alternatives = CurrenciesSupported(True, True)
97                IsHyphenated = True
98            Case "fixedcashflows|end date"
99                FixedCashflowsDoubleClickHandler Target, Target.Offset(, gCN_Notional1 - gCN_EndDate), "date"
100               Cancel = True
101               Exit Sub
102           Case "fixedcashflows|notional 1"
103               FixedCashflowsDoubleClickHandler Target.Offset(, -gCN_Notional1 + gCN_EndDate), Target, "amount"
104               Cancel = True
105               Exit Sub
106           Case "fxoptionstrip|end date", "fxforwardstrip|end date"
107               FxStripDoubleClickHandler TradeType, Target, Target.Offset(, gCN_Notional1 - gCN_EndDate), Target.Offset(, gCN_Notional2 - gCN_EndDate), "date"
108               Cancel = True
109               Exit Sub
110           Case "fxoptionstrip|notional 1", "fxforwardstrip|notional 1"
111               FxStripDoubleClickHandler TradeType, Target.Offset(, gCN_EndDate - gCN_Notional1), Target, Target.Offset(, gCN_Notional2 - gCN_Notional1), "amount"
112               Cancel = True
113               Exit Sub
114           Case "fxoptionstrip|notional 2", "fxforwardstrip|notional 2"
115               FxStripDoubleClickHandler TradeType, Target.Offset(, gCN_EndDate - gCN_Notional2), Target.Offset(, gCN_Notional1 - gCN_Notional2), Target, "amount"
116               Cancel = True
117               Exit Sub
118           Case "interestrateswap|notional 1", "interestrateswap|notional 2", "crosscurrencyswap|notional 1", "crosscurrencyswap|notional 2"
119               Alternatives = NotionalDoubleClickHandler(Target)
120               Cancel = True
121       End Select

122       BackUpRange EnlargedTarget, shUndo, , True
123       originalValue = EnlargedTarget.text

124       If Not IsEmpty(Alternatives) Then
125           Cancel = True
126           Select Case sNRows(Alternatives)
                  Case 1
127                   EnlargedTarget.Value = FirstElementOf(Alternatives)
128               Case 2
129                   If Target.Value = Alternatives(1, 1) Then
130                       EnlargedTarget.Value = Alternatives(2, 1)
131                   Else
132                       EnlargedTarget.Value = Alternatives(1, 1)
133                   End If
134               Case Is <= 8
135                   If IsHyphenated Then
136                       For Each V In Alternatives
137                           If sEquals(sStringBetweenStrings(V, , " - "), CStr(Target.Value)) Then
138                               CurrentChoice = V
139                           End If
140                       Next
141                       Res = ShowOptionButtonDialog(sArrayMakeText(Alternatives), MsgBoxTitle(), TopText, CurrentChoice, Target.Offset(0, 1), False, , , HelpMethodName)
142                       If Not IsEmpty(Res) Then
143                           Res = sStringBetweenStrings(Res, , " - ")
144                           If IsNumeric(Res) Then Res = CDbl(Res)
145                       End If
146                   Else
147                       Res = ShowOptionButtonDialog(sArrayMakeText(Alternatives), MsgBoxTitle(), TopText, CurrentChoice, Target.Offset(0, 1), False, , , HelpMethodName)
148                   End If

149                   If Not IsEmpty(Res) Then
150                       EnlargedTarget.Value = Res    'will morph strings "1" etc to numbers
151                   End If
152               Case Else
153                   If IsHyphenated Then
154                       For Each V In Alternatives
155                           If sEquals(sStringBetweenStrings(V, , " - "), CStr(Target.Value)) Then
156                               CurrentChoice = V
157                           End If
158                       Next
159                       Res = ShowSingleChoiceDialog(sArrayMakeText(Alternatives), , , , , MsgBoxTitle(), TopText, Target.Offset(0, 1), , gProjectName + TradeAttribute)
160                       If Not IsEmpty(Res) Then
161                           Res = sStringBetweenStrings(Res, , " - ")
162                           If IsNumeric(Res) Then Res = CDbl(Res)
163                       End If
164                   Else
165                       Res = ShowSingleChoiceDialog(sArrayMakeText(Alternatives), , , , , MsgBoxTitle(), TopText, Target.Offset(0, 1), , gProjectName + TradeAttribute)
166                   End If
167                   If Not IsEmpty(Res) Then
168                       EnlargedTarget.Value = Res
169                   End If
170           End Select
171           CalculatePortfolioSheet
172           Application.OnUndo "Restore " + Replace(EnlargedTarget.Address, "$", "") + IIf(IsNull(originalValue), " to its previous values", " to " + originalValue), "RestoreRange"
173       End If

174       Exit Sub
ErrHandler:
175       Throw "#TradeDoubleClickHandler (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub ShowHelpOnDayCounts()
          Dim Prompt As String
1         On Error GoTo ErrHandler
2         Prompt = sConcatenateStrings(shHiddenSheet.Range("Help_on_Day_Count_Types").Value, vbLf)
3         MsgBoxPlus Prompt, vbInformation, "Help on Day Count Types", , , , , 400
4         Exit Sub
ErrHandler:
5         Throw "#ShowHelpOnDayCounts (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

