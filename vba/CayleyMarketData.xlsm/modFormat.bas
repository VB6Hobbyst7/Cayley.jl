Attribute VB_Name = "modFormat"
Public Const gTextColor = 13395456
Option Explicit

Sub TestFormatCurrencySheet()
1         On Error GoTo ErrHandler
2         FormatCurrencySheet ActiveSheet, True, True
3         ActiveSheet.Protect
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestFormatCurrencySheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatCurrencySheet
' Author    : Philip Swannell
' Date      : 13-Jun-2016
' Purpose   : Apply cell formatting to the cells of a currency sheet
'             CollapseColumns can be True, False or empty (leave collapse state as is)
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatCurrencySheet(ws As Worksheet, ClearComments As Boolean, CollapseColumns As Variant)
          Dim b As Button
          Dim c As Range
          Dim Ccy As String
          Dim Collapsed As Boolean
          Dim CollateralCcy As String
          Dim Comment As String
          Dim i As Long
          Dim MaxCol As Long
          Dim MaxRow As Long
          Dim o As Object
          Dim RangeToAutoFit As Range
          Dim RangeToFormat As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TLSwaptionAddress

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(ws)
3         Set SUH = CreateScreenUpdateHandler()
4         Ccy = Left(ws.Name, 3)
5         CollateralCcy = shConfig.Range("CollateralCcy").Value

6         ws.Calculate

7         Set o = ws
8         o.UsedRange    'Resets the used range

9         For Each b In ws.Buttons
10            If InStr(b.OnAction, "GroupingButton") > 0 Then
11                If b.Caption = " >" Or b.Caption = " " & Chr(125) Then
12                    Collapsed = True
13                End If
14            End If
15        Next

16        GroupingButtonDoAllOnSheet ws, True

          'Correct title
17        With RangeFromSheet(ws, "Title")
18            .ClearFormats
19            .Formula = "=RIGHT(sCellInfo(""SheetName""),3)&"" curves and volatility"""
20            .Font.Size = 22
21        End With

          'ensure that all empty cells have no formats etc.
22        TLSwaptionAddress = RangeFromSheet(ws, "VolInit").Cells(0, 0).Address
23        For Each c In ws.UsedRange.Cells
24            If IsEmpty(c.Value) Then
25                If c.Address <> TLSwaptionAddress Then
26                    c.Clear
27                End If
28            Else
29                If c.Column > MaxCol Then
30                    MaxCol = c.Column
31                End If
32                If c.Row > MaxRow Then
33                    MaxRow = c.Row
34                End If
35            End If
36        Next c
37        With ws.UsedRange
38            If .Row + .Rows.Count - 1 > MaxRow Then
39                ws.Cells(MaxRow + 1, 1).Resize(.Row + .Rows.Count - 1 - MaxRow).EntireRow.Delete
40            End If
41            If .Column + .Columns.Count - 1 > MaxCol Then
42                ws.Cells(1, MaxCol + 1).Resize(, .Column + .Columns.Count - 1 - MaxCol).EntireColumn.Delete
43            End If
44            o.UsedRange    'Resets the used range
45        End With

46        ws.UsedRange.Locked = True

          Dim cnBB As Long
          Dim cnFixDCT As Long
          Dim cnFixFreq As Long
          Dim cnFloatDCT As Long
          Dim cnFloatFreq As Long
          Dim cnRate As Long
          Dim cnTenor As Long
47        Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "SwapRatesInit").Resize(, 6))
48        With RangeToFormat
49            If ClearComments Then
50                .ClearComments
51                .Interior.ColorIndex = xlColorIndexAutomatic
52            End If
53            .Font.Color = gTextColor
54            .Font.Name = "Calibri"
55            .Font.Size = 11
56            .Locked = False
57            .HorizontalAlignment = xlHAlignCenter
58            .VerticalAlignment = xlVAlignCenter
59            AddGreyBorders .Offset(0), True
60            AddGreyBorders .Resize(, 2), True
61            .NumberFormat = "General"
62            .Validation.Delete
63            For i = 1 To 7
64                Select Case .Cells(0, i).Value
                      Case "Tenor"
65                        cnTenor = i
66                    Case "Rate"
67                        cnRate = i
68                        .Columns(i).NumberFormat = "0.000%;[Red]-0.000%"
69                    Case "FixFreq"
70                        cnFixFreq = i
71                        SetValidationForFreq .Columns(i)
72                    Case "FixFreq", "FloatFreq"
73                        cnFloatFreq = i
74                        SetValidationForFreq .Columns(i)
75                    Case "FloatDCT"
76                        cnFloatDCT = i
77                        SetValidationForFloatDCT .Columns(i)
78                    Case "FixDCT"
79                        cnFixDCT = i
80                        SetValidationForFixedDCT .Columns(i)
81                    Case "BloombergCode"    'PGS 16/3/18 No longer have Bloomberg codes on the sheet, but keep this code in case we want to  put them back
82                        cnBB = i
83                        With .Columns(i)
84                            Comment = "Bloomberg ticker shown for information only," + vbLf + _
                                  "as returned by VBA function" + vbLf + _
                                  "BloombergTickerInterestRateSwap"
85                            SetCellComment .Cells(1, 1), Comment
86                            .FormulaR1C1 = "=BloombergTickerInterestRateSwap(RIGHT(sCellInfo(""SheetName""),3),RC[" + CStr(cnTenor - cnBB) + "],RC[" + CStr(cnFixFreq - cnBB) + "],RC[" + CStr(cnFloatFreq - cnBB) + "])"
87                            .Value = .Value
88                            .Font.ColorIndex = xlColorIndexAutomatic
89                            .Locked = True
90                        End With
91                End Select
92            Next i
93        End With

94        Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "XccyBasisSpreadsInit"))
95        With RangeToFormat
96            If ClearComments Then
97                .ClearComments
98                .Interior.ColorIndex = xlColorIndexAutomatic
99            End If
100           .Font.Color = gTextColor
101           .Font.Name = "Calibri"
102           .Font.Size = 11
103           .Locked = False
104           .HorizontalAlignment = xlHAlignCenter
105           .VerticalAlignment = xlVAlignCenter
106           AddGreyBorders .Offset(0), True
107           AddGreyBorders .Resize(, 2), True
108           .NumberFormat = "General"
109           .Columns(2).NumberFormat = "0.000%;[Red]-0.000%"
110           .Validation.Delete
111           SetValidationForFreq .Columns(3)
112           SetValidationForFixedDCT .Columns(4)
113           SetValidationForFreq .Columns(5)
114           SetValidationForFixedDCT .Columns(6)
115           If False Then    'PGS 16/3/18 No longer have Bloomberg codes on the sheet, but keep this code in case we want to  put them back
116               With .Columns(7)
117                   Comment = "Bloomberg ticker shown for information only," + vbLf + _
                          "as returned by VBA function" + vbLf + _
                          "BloombergTickerBasisSwap"
118                   SetCellComment .Cells(1, 1), Comment
119                   .FormulaR1C1 = "=BloombergTickerBasisSwap(sCellInfo(""SheetName""),RC[-6])"
120                   .Value = .Value
121                   .Font.ColorIndex = xlColorIndexAutomatic
122                   .Locked = True
123               End With
124           End If
125       End With

126       With ws.Range("Spread_is_on")
127           SetValidation .Offset(0), Ccy & "," & CollateralCcy, ""
128           .Locked = False
129           .Font.Color = gTextColor
130       End With

131       Set RangeToAutoFit = Application.Union(sExpandDown(RangeFromSheet(ws, "XccyBasisSpreadsInit").Rows(0)), _
              sExpandDown(RangeFromSheet(ws, "SwapRatesInit").Rows(-1).Resize(1, 7)))

132       RangeToAutoFit.Columns.AutoFit

133       Set RangeToFormat = sexpandRightDown(RangeFromSheet(ws, "VolInit").Cells(0, 0))

134       With RangeToFormat
135           If ClearComments Then
136               .ClearComments
137               .Interior.ColorIndex = xlColorIndexAutomatic
138           End If
139           .Font.Name = "Calibri"
140           .Font.Size = 11
141           .HorizontalAlignment = xlHAlignCenter
142           .VerticalAlignment = xlVAlignCenter
143           AddGreyBorders .Offset(0), True
144           .NumberFormat = "General"

145           .Locked = True
146           With .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1)
147               .Font.Color = gTextColor
148               .NumberFormat = "0.00%"
149               .Locked = False
150           End With
151           .Columns.AutoFit
152           For i = 2 To .Columns.Count
153               .Columns(i).ColumnWidth = .Columns(i).ColumnWidth + 1
154           Next
155       End With

156       Set RangeToFormat = RangeFromSheet(ws, "SwaptionVolParameters")
157       With RangeToFormat
158           .ClearFormats
159           .Columns(1).HorizontalAlignment = xlHAlignRight
160           .Columns(2).HorizontalAlignment = xlHAlignLeft
161           AutoFitColumns .Offset(0), 1
162           AddGreyBorders .Offset(0), True
163           For Each c In .Columns(2).Cells
164               Select Case c.Offset(0, -1).Value
                      Case "FixedFrequency", "FloatingFrequency"
165                       SetValidationForFreq c
166                       c.Font.Color = gTextColor
167                       .Locked = False
168                   Case "FixedDCT"
169                       SetValidationForFixedDCT c
170                       c.Font.Color = gTextColor
171                       .Locked = False
172                   Case "FloatingDCT"
173                       SetValidationForFloatDCT c
174                       c.Font.Color = gTextColor
175                       .Locked = False
176                   Case "QuoteType"
177                       SetValidation c, "Normal,Log Normal,OIS Normal,OIS Log Normal", "That's not a valid QuoteType"
178                       c.Font.Color = gTextColor
179                       .Locked = False
180                   Case "Contributor"
181                       SetValidation c, "BBIR,CFIR,CMPL,CMPN,CNTR,GFIS,ICPL,LAST,SMKO,TRPU", "That's not a valid Contributor"
182                       c.Font.Color = gTextColor
183                       .Locked = False
184                   Case "example Code"
185                       .Locked = True
186               End Select
187           Next c
188       End With

          'For Libor Transition
          'In case worksheet has not yet been amended for LiborTransition
189       If Not IsInCollection(ws.Names, "FloatingLegType") Then
              Dim Target As Range

190           With RangeFromSheet(ws, "SwaptionVolParameters")
191               Set Target = .Cells(.Rows.Count + 2, 1)
192           End With

193           With Target
194               .ClearFormats
195               .Value = "Libor Transition"
196           End With

197           With Target.Offset(1)
198               .ClearFormats
199               .Value = "FloatingLegType"
200           End With

201           With Target.Offset(1, 1)
202               .Clear
203               .Value = IIf(ws.Name = "EUR", "IBOR", "RFR")
204               ws.Names.Add "FloatingLegType", .Cells
205           End With
206       End If

207       With RangeFromSheet(ws, "FloatingLegType")
208           .HorizontalAlignment = xlHAlignLeft
209           .Locked = False
210           .Font.Color = gTextColor
211           With .Validation
                .Delete
212               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                      xlBetween, Formula1:="RFR,IBOR"
213           End With
           .Offset(0, -1).HorizontalAlignment = xlHAlignRight
           AddGreyBorders .Offset(0, -1).Resize(1, 2), True
           
214       End With




          'Set width of empty columns
215       For Each c In ws.UsedRange.Rows(1).Cells
216           If IsEmpty(c.Value) Then
217               If c.End(xlDown).Row = ws.Rows.Count Then
218                   c.ColumnWidth = 2
219               End If
220           End If
221       Next

222       If VarType(CollapseColumns) = vbBoolean Then
223           GroupingButtonDoAllOnSheet ws, Not (CollapseColumns)
224       Else
225           GroupingButtonDoAllOnSheet ws, Not (Collapsed)
226       End If

227       For Each b In ws.Buttons
228           If InStr(b.OnAction, "ShowMenu") > 0 Then

229               b.Top = 3
230               b.Left = 247
231               b.Height = 24
232               b.Width = 65
233               b.Placement = xlFreeFloating
234               b.Caption = "Menu..."
235           End If
236       Next

237       Exit Sub
ErrHandler:
238       Throw "#FormatCurrencySheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AmendInflationSheets
' Author    : Philip Swannell
' Date      : 25-Apr-2017
' Purpose   : Ad hoc method
' -----------------------------------------------------------------------------------------------------------------------
Sub AmendInflationSheets()
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             If IsInflationSheet(ws) Then
4                 FormatInflationSheet ws, False, False
5                 ws.Protect , True, True
6             End If
7         Next
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#AmendInflationSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatInflationSheet
' Author    : Philip Swannell
' Date      : 13-Jun-2016
' Purpose   : Apply cell formatting to the cells of an inflation sheet
'             CollapseColumns can be True, False or empty (leave collapse state as is)
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatInflationSheet(ws As Worksheet, ClearComments As Boolean, CollapseColumns As Variant)
          Dim b As Button
          Dim c As Range
          Dim Collapsed As Boolean
          Dim o As Object
          Dim RangeToFormat As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(ws)
3         Set SUH = CreateScreenUpdateHandler()

4         ws.Calculate

5         Set o = ws
6         o.UsedRange    'Resets the used range
7         ws.UsedRange.ClearFormats

8         For Each b In ws.Buttons
9             If InStr(b.OnAction, "GroupingButton") > 0 Then
10                If b.Caption = " >" Or b.Caption = " " & Chr(125) Then
11                    Collapsed = True
12                End If
13            End If
14        Next

15        GroupingButtonDoAllOnSheet ws, True

          'Correct title
16        With RangeFromSheet(ws, "Title")
17            .ClearFormats
18            .Formula = "=""Inflation: ""&sCellInfo(""SheetName"")"
19            .Font.Size = 22
20        End With

21        ws.UsedRange.Locked = True

22        Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "ZCSwapsInit")).Resize(, 3)
23        With RangeToFormat
24            .Cells(0, 1).Value = "Tenor": .Cells(0, 2).Value = "Rate": .Cells(0, 3).Value = "BloombergCode"

25            If ClearComments Then
26                .ClearComments
27                .Interior.ColorIndex = xlColorIndexAutomatic
28            End If
29            .ClearFormats
30            .NumberFormat = "General"
31            .HorizontalAlignment = xlHAlignCenter
32            .VerticalAlignment = xlVAlignCenter
33            .Font.Name = "Calibri"
34            .Font.Size = 11
35            .Font.Bold = False
36            .Font.Italic = False
37            With .Columns(2)
38                .NumberFormat = "0.000%;[Red]-0.000%"
39                .Locked = False
40                .Font.Color = gTextColor
41            End With
42            AddGreyBorders .Resize, True
43            With .Rows(0)
44                .Font.Bold = True
45                .HorizontalAlignment = xlHAlignCenter
46            End With
47            With .Columns(3)
48                .FormulaR1C1 = "=BloombergTickerZCInflationSwap(sCellInfo(""SheetName""),RC[-2])"
49                .Value = .Value
50            End With
51            AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 3
52        End With

53        Set RangeToFormat = RangeFromSheet(ws, "SeasonalAdjustments")
54        With RangeToFormat
55            .Cells(0, 1).Value = "Month": .Cells(0, 2).Value = "Adj"
56            .Columns(1).Value = sTokeniseString("'Jan,'Feb,'Mar,'Apr,'May,'Jun,'Jul,'Aug,'Sep,'Oct,'Nov,'Dec")
57            If ClearComments Then
58                .ClearComments
59                .Interior.ColorIndex = xlColorIndexAutomatic
60            End If
61            .ClearFormats
62            .NumberFormat = "General"
63            .HorizontalAlignment = xlHAlignCenter
64            .VerticalAlignment = xlVAlignCenter
65            .Font.Name = "Calibri"
66            .Font.Size = 11
67            .Font.Bold = False
68            .Font.Italic = False
69            With .Columns(2)
70                .NumberFormat = "General"
71                .Locked = False
72                .Font.Color = gTextColor
73            End With
74            AddGreyBorders .Resize, True
75            With .Rows(0)
76                .Font.Bold = True
77                .HorizontalAlignment = xlHAlignCenter
78            End With
79            AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 5
80        End With

81        Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "HistoricDataInit"))
82        With RangeToFormat
83            .Cells(0, 1).Value = "Year": .Cells(0, 2).Value = "Month": .Cells(0, 3).Value = "Index"
84            If ClearComments Then
85                .ClearComments
86                .Interior.ColorIndex = xlColorIndexAutomatic
87            End If
88            .ClearFormats
89            .NumberFormat = "General"
90            .HorizontalAlignment = xlHAlignCenter
91            .VerticalAlignment = xlVAlignCenter
92            .Font.Name = "Calibri"
93            .Font.Size = 11
94            .Font.Bold = False
95            .Font.Italic = False
96            AddGreyBorders .Resize, True
97            With .Rows(0)
98                .Font.Bold = True
99                .HorizontalAlignment = xlHAlignCenter
100           End With
101           AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 3
102       End With

103       Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "ParametersInit"))
104       With RangeToFormat
105           ws.Names.Add "Parameters", RangeToFormat
106           If ClearComments Then
107               .ClearComments
108               .Interior.ColorIndex = xlColorIndexAutomatic
109           End If
110           .ClearFormats
111           .NumberFormat = "General"
112           .Columns(1).HorizontalAlignment = xlHAlignRight
113           .Columns(2).HorizontalAlignment = xlHAlignLeft
114           .VerticalAlignment = xlVAlignCenter
115           .Font.Name = "Calibri"
116           .Font.Size = 11
117           .Font.Bold = False
118           .Font.Italic = False
119           AddGreyBorders .Resize, True
120           AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 2
121           .Validation.Delete
122           For Each c In .Columns(2).Cells
123               Select Case c.Offset(0, -1).Value
                  Case "LagMethod"
124                   SetValidation c, sConcatenateStrings(sSupportedInflationLagMethods()), "Please select a valid LagMethod. Use in-cell dropdown to see allowed values."
125                   c.Locked = False
126                   c.Font.Color = gTextColor
127               Case "effectiveDate"
128                   SetValidation c, sConcatenateStrings(sSupportedInflationeffectiveDates()), "Please select a valid effective Date. Use in-cell dropdown to see allowed values."
129                   c.Locked = False
130                   c.Font.Color = gTextColor
131               End Select
132           Next c
133       End With

134       Set RangeToFormat = sExpandDown(RangeFromSheet(ws, "InflationVolInit"))
135       With RangeToFormat
136           If ClearComments Then
137               .ClearComments
138               .Interior.ColorIndex = xlColorIndexAutomatic
139           End If
140           .ClearFormats
141           .NumberFormat = "General"
142           .HorizontalAlignment = xlHAlignCenter
143           .VerticalAlignment = xlVAlignCenter
144           .Font.Name = "Calibri"
145           .Font.Size = 11
146           .Font.Bold = False
147           .Font.Italic = False
148           AddGreyBorders .Resize, True
149           With .Rows(0)
150               .Font.Bold = True
151               .HorizontalAlignment = xlHAlignCenter
152           End With
153           AutoFitColumns .Offset(-1).Resize(.Rows.Count + 1), 3
154       End With
          'Set width of empty columns
155       For Each c In ws.UsedRange.Rows(1).Cells
156           If IsEmpty(c.Value) Then
157               If c.End(xlDown).Row = ws.Rows.Count Then
158                   c.ColumnWidth = 2
159               End If
160           End If
161       Next

162       If VarType(CollapseColumns) = vbBoolean Then
163           GroupingButtonDoAllOnSheet ws, Not (CollapseColumns)
164       Else
165           GroupingButtonDoAllOnSheet ws, Not (Collapsed)
166       End If

167       For Each b In ws.Buttons
168           If InStr(b.OnAction, "ShowMenu") > 0 Then
169               With ws.Cells(2, 2)
170                   b.Top = .Top
171                   b.Left = .Left
172                   b.Height = .Height
173                   b.Width = .Width
174                   b.Placement = xlMove
175                   b.Caption = "Menu..."
176               End With
177           End If
178       Next

179       Exit Sub
ErrHandler:
180       Throw "#FormatInflationSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetValidation
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Set the validation on a range to an arbitrary list
' -----------------------------------------------------------------------------------------------------------------------
Private Function SetValidation(R As Range, List As String, errorMessage As String)
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
3         With R.Validation
4             .Delete
5             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                   Operator:=xlBetween, Formula1:=List
6             .IgnoreBlank = True
7             .InCellDropdown = True
8             .InputTitle = ""
9             .ErrorTitle = "Cayley Market Data"
10            .InputMessage = ""
11            .errorMessage = errorMessage
12            .ShowInput = True
13            .ShowError = True
14        End With

15        Exit Function
ErrHandler:
16        Throw "#SetValidation (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetValidationForFreq
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Set the validation for a cell to what it should be for a coupon frequency
' -----------------------------------------------------------------------------------------------------------------------
Private Function SetValidationForFreq(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                   Operator:=xlBetween, Formula1:="Ann,Semi,Quart"
5             .IgnoreBlank = True
6             .InCellDropdown = True
7             .InputTitle = ""
8             .ErrorTitle = "Cayley Market Data"
9             .InputMessage = ""
10            .errorMessage = _
              "That's not a valid payment frequency."
11            .ShowInput = True
12            .ShowError = True
13        End With

14        Exit Function
ErrHandler:
15        Throw "#SetValidationForFreq (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetValidationForFloatDCT
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Set the validation for a cell to what it should be for a floating daycount
' -----------------------------------------------------------------------------------------------------------------------
Private Function SetValidationForFloatDCT(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                   Operator:=xlBetween, Formula1:="A/360,A/365F"
5             .IgnoreBlank = True
6             .InCellDropdown = True
7             .InputTitle = ""
8             .ErrorTitle = "Cayley Market Data"
9             .InputMessage = ""
10            .errorMessage = _
              "That's not a valid floating daycount."
11            .ShowInput = True
12            .ShowError = True
13        End With

14        Exit Function
ErrHandler:
15        Throw "#SetValidationForFloatDCT (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetValidationForFixedDCT
' Author    : Philip Swannell8
' Date      : 15-Jun-2016
' Purpose   : Set the validation for a cell to what it should be for a fixed daycount
' -----------------------------------------------------------------------------------------------------------------------
Private Function SetValidationForFixedDCT(R As Range)
1         On Error GoTo ErrHandler
2         With R.Validation
3             .Delete
4             .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                   Operator:=xlBetween, Formula1:="30/360,30e/360,30e/360 (ISDA),A/360,A/365F,Act/Act"
5             .IgnoreBlank = True
6             .InCellDropdown = True
7             .InputTitle = ""
8             .ErrorTitle = "Cayley Market Data"
9             .InputMessage = ""
10            .errorMessage = _
              "That's not a valid fixed daycount."
11            .ShowInput = True
12            .ShowError = True
13        End With

14        Exit Function
ErrHandler:
15        Throw "#SetValidationForFixedDCT (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub TestFormatFxVolSheet()
1         On Error GoTo ErrHandler
2         FormatFxVolSheet False
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestFormatFxVolSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatFxVolSheet
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Applies cell formatting to the FX sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatFxVolSheet(ClearComments As Boolean)

          Dim Colors As Variant
          Dim i As Long
          Dim j As Long
          Dim ls As String
          Dim Numeraire As String
          Dim RangeNoHeaders As Range
          Dim RangeWithHeaders As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         ls = Application.International(xlListSeparator)

3         Set ws = ThisWorkbook.Worksheets("FX")

4         Set SPH = CreateSheetProtectionHandler(ws)
5         Set SUH = CreateScreenUpdateHandler()

6         Set RangeWithHeaders = sexpandRightDown(RangeFromSheet(ws, "FxDataTopLeft"))
7         With RangeWithHeaders
8             Set RangeNoHeaders = .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1)
9         End With

10        With RangeWithHeaders
11            If ClearComments Then
12                .ClearComments
13                .ClearFormats
14            Else
15                Colors = sReshape(0, .Rows.Count, .Columns.Count)
16                For i = 1 To .Rows.Count
17                    For j = 1 To .Columns.Count
18                        Colors(i, j) = .Cells(i, j).Interior.Color
19                    Next j
20                Next i
21                .ClearFormats
22                For i = 1 To .Rows.Count
23                    For j = 1 To .Columns.Count
24                        .Cells(i, j).Interior.Color = Colors(i, j)
25                    Next j
26                Next i
27            End If
28            .HorizontalAlignment = xlHAlignCenter
29            AddGreyBorders .Offset(0), True
30        End With

31        If RangeWithHeaders.Row + RangeWithHeaders.Rows.Count < shFx.UsedRange.Row + shFx.UsedRange.Rows.Count Then
32            With RangeWithHeaders
33                .Cells(.Rows.Count + 1, 1).Resize(shFx.UsedRange.Row + shFx.UsedRange.Rows.Count - .Row - .Rows.Count).EntireRow.Delete
34            End With
              Dim o As Object
35            Set o = shFx
36            o.UsedRange
37        End If

38        With RangeNoHeaders
39            .Locked = False
40            .Font.Color = gTextColor
41            .Columns(1).NumberFormat = "[>=100]#,##0.00;[>=10]#,##0.000;#,##0.0000"
42            .ColumnWidth = 8
43            With .Offset(, 1).Resize(, .Columns.Count - 1)
44                .NumberFormat = "0.00%"
45            End With
46            With .Offset(0, -1).Resize(, .Columns.Count + 1)
                  Dim CellAddress As String
                  Dim Formula1 As String
47                Numeraire = RangeFromSheet(shConfig, "Numeraire", False, True, False, False, False).Value
48                CellAddress = "$" & Replace(.Cells(1, 1).Address, "$", "")
49                Formula1 = "=ISeRROR(FIND(""" + Numeraire + """" & ls & CellAddress + "))"
50                .FormatConditions.Delete
51                .FormatConditions.Add Type:=xlExpression, Formula1:=Formula1
52                .FormatConditions(1).SetFirstPriority
53                With .FormatConditions(1).Font
54                    .ThemeColor = xlThemeColorDark1
55                    .TintAndShade = -0.349986266670736
56                End With
57            End With
58        End With

59        With sexpandRightDown(RangeFromSheet(ws, "HistoricFxVolsTopLeft"))
60            .ClearFormats
61            With .Offset(, 1).Resize(, .Columns.Count - 1)
62                .NumberFormat = "0.00%"
63                .Font.Color = gTextColor
64                .Locked = False
65            End With
66            .HorizontalAlignment = xlHAlignCenter
67            AddGreyBorders .Offset(0), True
68            .ColumnWidth = 8
69            Numeraire = RangeFromSheet(shConfig, "Numeraire", False, True, False, False, False).Value
70            CellAddress = "$" & Replace(.Cells(1, 1).Address, "$", "")
71            Formula1 = "=ISeRROR(FIND(""" + Numeraire + """" + ls + CellAddress + "))"
72            .FormatConditions.Delete
73            .FormatConditions.Add Type:=xlExpression, Formula1:=Formula1
74            .FormatConditions(1).SetFirstPriority
75            With .FormatConditions(1).Font
76                .ThemeColor = xlThemeColorDark1
77                .TintAndShade = -0.349986266670736
78            End With
79        End With

80        Set SPH = Nothing
81        shFx.Protect

82        Exit Sub
ErrHandler:
83        Throw "#FormatFxVolSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatCreditSheet2
' Author    : Philip Swannell
' Date      : 16-Jan-2017
' Purpose   : Wrapper to FormatCreditSheet with error handling that will work via application.run
' -----------------------------------------------------------------------------------------------------------------------
Function FormatCreditSheet2()
1         On Error GoTo ErrHandler
2         FormatCreditSheet False
3         Exit Function
ErrHandler:
4         FormatCreditSheet2 = "#FormatCreditSheet2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatCreditSheet
' Author    : Hermione Glyn
' Date      : 16-Jan-2017
' Purpose   : Formats the CDS data and tickers. For use when adding a new couterparty
'             to the Credit sheet (AddBankToCreditSheet -> AddCreditCounterparty).
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatCreditSheet(ClearComments As Boolean)
          Dim CDSRangeNoHeaders As Range
          Dim CDSRangeWithHeaders As Range
          Dim Colors As Variant
          Dim HeaderRows As Range
          Dim i As Long
          Dim j As Long
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TickersWithHeaders As Range
          Dim ws As Worksheet
1         On Error GoTo ErrHandler

2         Set ws = shCredit
3         Set SPH = CreateSheetProtectionHandler(ws)
4         Set SUH = CreateScreenUpdateHandler()

5         Set HeaderRows = sexpandRightDown(RangeFromSheet(ws, "CDSTopLeft")).Offset(-1, 0).Resize(2, 19)
6         Set CDSRangeWithHeaders = CDSRange(ws)
7         With CDSRangeWithHeaders
8             Set CDSRangeNoHeaders = .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1)
9         End With
10        Set TickersWithHeaders = sexpandRightDown(RangeFromSheet(ws, "CDSTickersTopLeft")).Resize(CDSRangeWithHeaders.Rows.Count)

11        With CDSRangeWithHeaders
12            If ClearComments Then
13                .ClearComments
14                .ClearFormats
15            Else
16                Colors = sReshape(0, .Rows.Count, .Columns.Count)
17                For i = 1 To .Rows.Count
18                    For j = 1 To .Columns.Count
19                        Colors(i, j) = .Cells(i, j).Interior.Color
20                    Next j
21                Next i
22                .ClearFormats
23                For i = 1 To .Rows.Count
24                    For j = 1 To .Columns.Count
25                        .Cells(i, j).Interior.Color = Colors(i, j)
26                    Next j
27                Next i
28            End If
29            .HorizontalAlignment = xlHAlignCenter
30        End With

31        With CDSRangeNoHeaders
32            .Locked = False
33            .Font.Color = gTextColor
34            .ColumnWidth = 8
35            .NumberFormat = "0.00%"
36            .Columns(2).NumberFormat = "0%"
37        End With

38        With TickersWithHeaders
39            .ClearFormats
40            .ColumnWidth = 10
41        End With

42        AddGreyBorders CDSRangeNoHeaders, True
43        AddGreyBorders CDSRangeNoHeaders.Resize(, 2), True
44        AddGreyBorders CDSRangeNoHeaders.Columns(0), True

45        AddGreyBorders TickersWithHeaders, True
46        AddGreyBorders HeaderRows, True
47        AddGreyBorders CDSRange(shCredit).Offset(-1).Resize(2), True
48        AddGreyBorders HeaderRows.Columns(1), True
49        AddGreyBorders CDSRangeNoHeaders.Offset(-2).Resize(2, 2), True
50        Exit Sub
ErrHandler:
51        Throw "#FormatCreditSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

