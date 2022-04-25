VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shCreditUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_BeforeDoubleClick
' Author    : Philip Swannell
' Date      : 03-Aug-2016
' Purpose   : Double-click handling...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim Alternatives As Variant
          Dim AttributeName As String
          Dim Res As Variant
          Dim UseRegEx As Boolean
          Const chAdvanced = "Advanced filtering..."
          Const chSimple = "Pick from List..."
          Const MAXMRUsToShow = 9
          Dim AmPickingBank As Boolean
          Dim AnchorObject As Object
          Dim EnableFlags
          Dim FaceIDs
          Dim Filters
          Dim IncludeFutureTrades As Boolean
          Dim LinesBook As Workbook
          Dim LinesBookIsOpen As Boolean
          Dim MarketBookIsOpen As Boolean
          Dim PortfolioAgeing As Double
          Dim RegKey As String
          Dim TheChoices
          Dim TradesBook As Workbook
          Dim TradesBookIsOpen As Boolean

1         On Error GoTo ErrHandler

2         If sElapsedTime() - LastAltBacktickTime < 0.5 Then
3             Set AnchorObject = Target.offset(0, 1)
4         End If

5         RunThisAtTopOfCallStack

6         If Not OtherBooksAreOpen(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen) Then
7             PleaseOpenOtherBooks
8         End If

9         Set LinesBook = OpenLinesWorkbook(True, False)
10        Set TradesBook = OpenTradesWorkbook(True, False)

11        Cancel = True
12        Select Case Target.Address
              Case SafeRangeAddress(Me, "NumMCPaths")
13                Alternatives = sArrayStack(127, 255, 511, 1023, 2047)
14            Case SafeRangeAddress(Me, "NumObservations")
15                Alternatives = sArrayStack(10, 20, 40, 60, 80, 100, 160, 320)
16            Case SafeRangeAddress(Me, "FilterBy1"), SafeRangeAddress(Me, "FilterBy2")

17                If Target.Address = SafeRangeAddress(Me, "FilterBy1") Then
18                    RegKey = "CayleyFilterBy1_2022"
19                    AddFilterToMRU RegKey, "Counterparty Parent"

20                Else
21                    RegKey = "CayleyFilterBy2_2022"
22                    AddFilterToMRU RegKey, "None"
23                End If

24                GetMRUFilters RegKey, Filters, TheChoices, FaceIDs, EnableFlags
25                If Not IsMissing(Filters) Then
26                    If sNRows(Filters) > MAXMRUsToShow Then
27                        TheChoices = sSubArray(TheChoices, 1, 1, MAXMRUsToShow)
28                        FaceIDs = sSubArray(FaceIDs, 1, 1, MAXMRUsToShow)
29                        EnableFlags = sSubArray(EnableFlags, 1, 1, MAXMRUsToShow)
30                    End If
31                End If
32                TheChoices = sArrayStack(TheChoices, "--&" & chSimple)
33                FaceIDs = sArrayStack(FaceIDs, 447)
34                EnableFlags = sArrayStack(EnableFlags, True)
35                Res = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , AnchorObject, False)
36                Select Case Res
                      Case "#Cancel!"
37                        GoTo EarlyExit
38                    Case chSimple
39                        Alternatives = GetColumnFromTradesWorkbook("AllHeaders", False, 0, True, True, TradesBook, shFutureTrades, Date)
40                        Alternatives = sSortedArray(Alternatives)
41                        Alternatives = sArrayStack("None", Alternatives)
42                    Case Else
43                        SafeSetCellValue2 Target, CStr(Res)
44                        GoTo EarlyExit
45                End Select

46            Case SafeRangeAddress(Me, "Filter1Value"), SafeRangeAddress(Me, "Filter2Value")
                  Dim FilterBy As String
47                If Target.Address = SafeRangeAddress(Me, "Filter1Value") Then
48                    FilterBy = RangeFromSheet(Me, "FilterBy1").Value2
49                Else
50                    FilterBy = RangeFromSheet(Me, "FilterBy2").Value2
51                End If
52                AmPickingBank = UCase(CStr(FilterBy)) = UCase("Counterparty Parent")
53                If (LCase(FilterBy)) = "none" Then
54                    Target.Value = "None"
55                    GoTo EarlyExit
56                Else

57                    RegKey = "CayleyFilterBy_2022_" & FilterBy
58                    GetMRUFilters RegKey, Filters, TheChoices, FaceIDs, EnableFlags
59                    If Not IsMissing(Filters) Then
60                        If sNRows(Filters) > MAXMRUsToShow Then
61                            TheChoices = sSubArray(TheChoices, 1, 1, MAXMRUsToShow)

62                            FaceIDs = sSubArray(FaceIDs, 1, 1, MAXMRUsToShow)
63                            EnableFlags = sSubArray(EnableFlags, 1, 1, MAXMRUsToShow)
64                        End If
65                    End If
66                    TheChoices = sArrayStack(TheChoices, "--&" & chSimple, "&" & chAdvanced)
67                    If AmPickingBank Then TheChoices = AnnotateBankNames(TheChoices, True, LinesBook, True)
68                    FaceIDs = sArrayStack(FaceIDs, 447, 502)
69                    EnableFlags = sArrayStack(EnableFlags, True, True)
                      Dim PopUpRes
                      'The "show command bar then show full list is the use clicks 'Pick One'" paradigm does not work well for picking banks
                      'so we suppress it. PGS 22 March 2022
70                    If AmPickingBank Then
71                        PopUpRes = chSimple
72                    Else
73                        PopUpRes = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , AnchorObject, False)
74                    End If
75                    Select Case PopUpRes
                          Case "#Cancel!"
76                            GoTo EarlyExit
77                        Case chSimple, chAdvanced
78                            PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False)
79                            IncludeFutureTrades = RangeFromSheet(shCreditUsage, "IncludeFutureTrades", False, False, True, False, False)
80                            If AmPickingBank Then
81                                Alternatives = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
82                                Alternatives = sSortedArray(Alternatives)
83                                Alternatives = AnnotateBankNames(Alternatives, True, LinesBook, False)
84                            Else
85                                Alternatives = GetColumnFromTradesWorkbook(FilterBy, IncludeFutureTrades, PortfolioAgeing, True, True, TradesBook, shFutureTrades, Date)
86                            End If
87                            Alternatives = sRemoveDuplicates(Alternatives, True)
88                            UseRegEx = (PopUpRes = chAdvanced)
89                        Case Else
90                            If AmPickingBank Then PopUpRes = AnnotateBankNames(PopUpRes, False, LinesBook, True)
91                            SafeSetCellValue2 Target, CStr(PopUpRes)
92                            GoTo EarlyExit

93                    End Select
94                    Select Case UCase(FilterBy)
                          Case "Counterparty Parent"
95                            If PopUpRes = chSimple Then
96                                Alternatives = AnnotateBankNames(Alternatives, True, LinesBook)
97                                Res = ShowSingleChoiceDialog(Alternatives, , , , , , , AnchorObject, , "PFE Calculator")
98                                If IsEmpty(Res) Then
99                                    GoTo EarlyExit
100                               Else
101                                   Target.Value = AnnotateBankNames(Res, False, LinesBook)
102                                   GoTo EarlyExit
103                               End If
104                           End If
105                   End Select
106               End If
107           Case SafeRangeAddress(Me, "ExtraTradesAre")
108               Alternatives = AllowedExtraTradesAre(True)

109           Case SafeRangeAddress(Me, "IncludeExtraTrades"), SafeRangeAddress(Me, "IncludeFutureTrades")
110               Alternatives = sArrayStack("TRUE", "FALSE")
111               Res = ShowCommandBarPopup(Alternatives, , , UCase(CStr(Target.Value)), AnchorObject)
112               If Res <> "#Cancel!" Then
113                   SafeSetCellValue2 Target, CStr(Res)
114               End If
115               GoTo EarlyExit
116           Case SafeRangeAddress(Me, "LinesScaleFactor")
117               Alternatives = sArrayStack(1, 1.2, 1.4, 1.6, 1.8, 2)
118               Res = ShowCommandBarPopup(Alternatives, , , UCase(CStr(Target.Value)), AnchorObject)
119               If Res <> "#Cancel!" Then
120                   SafeSetCellValue2 Target, CStr(Res)
121               End If
122               GoTo EarlyExit
123           Case SafeRangeAddress(Me, "TradesScaleFactor")
                  Const Scale0 As String = "0 (Like cancelling all trades)"
                  Const ScaleM1 As String = "-1 (Flips trade direction, so shows PFE from Airbus perspective)"
124               Alternatives = sArrayStack(ScaleM1, Scale0, 0.5, 1, 1.5, 2)
125               Res = ShowCommandBarPopup(Alternatives, , , UCase(CStr(Target.Value)), AnchorObject)
126               If Res = Scale0 Then Res = 0
127               If Res = ScaleM1 Then Res = -1
128               If Res <> "#Cancel!" Then
129                   SafeSetCellValue2 Target, CStr(Res)
130               End If
131               GoTo EarlyExit
132           Case SafeRangeAddress(Me, "FxShock"), SafeRangeAddress(Me, "FxVolShock")
133               Alternatives = sArrayStack(0.8, 0.9, 1, 1.1, 1.2)
134           Case SafeRangeAddress(Me, "PortfolioAgeing")
135               Alternatives = sArrayStack(0, 0.25, 0.5, 0.75, 1, 2, 3, 4, 5)
136           Case SafeRangeAddress(Me, "IncludeAssetClasses")
137               Alternatives = sArrayStack("Rates and Fx", "Fx", "Rates")
138           Case Else
139               Cancel = False
140               GoTo EarlyExit
141       End Select

142       If UseRegEx Then
143           If Target.Address = RangeFromSheet(Me, "Filter2Value").Address Then
144               AttributeName = RangeFromSheet(Me, "FilterBy2").Value
145           Else
146               AttributeName = RangeFromSheet(Me, "FilterBy1").Value
147           End If

148           Res = ShowRegularExpressionDialog(IIf(CStr(Target.Value) = "None", "", Target.Value), _
                  AttributeName, Alternatives, AnchorObject, "Filter Trades", _
                  "Include trades for which:", False, RegKey, True)
149           If Res = "#User Cancel!" Then GoTo EarlyExit
150           SafeSetCellValue2 Target, CStr(Res)
151           GoTo EarlyExit
152       End If

153       If sNRows(Alternatives) < 20 Then
154           Res = ShowCommandBarPopup(Alternatives, , , CStr(Target.Value), AnchorObject)
155           If Res = "#Cancel!" Then GoTo EarlyExit
156           If AmPickingBank Then Res = AnnotateBankNames(Res, False, LinesBook, False)
157           SafeSetCellValue2 Target, CStr(Res)
158           If RegKey <> "" Then
159               AddFilterToMRU RegKey, CStr(Res)
160           End If

161           Application.GoTo Target

162       Else
163           Alternatives = sArrayMakeText(Alternatives)
164           Res = ShowSingleChoiceDialog(Alternatives, , , , , , , Target.offset(0, 1), , "PFE Calculator")
165           If Not IsEmpty(Res) Then
166               If AmPickingBank Then Res = AnnotateBankNames(Res, False, LinesBook, False)

167               SafeSetCellValue2 Target, CStr(Res)
168               If RegKey <> "" Then
169                   AddFilterToMRU RegKey, CStr(Res)
170               End If

171           End If
172       End If

EarlyExit:
173       Set LinesBook = Nothing
174       AddFilters

175       Exit Sub
ErrHandler:
176       SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!", vbInformation, "Worksheet_BeforeDoubleClick"
End Sub

Function SafeRangeAddress(ws As Worksheet, RangeName As String)
1         On Error GoTo ErrHandler
2         SafeRangeAddress = RangeFromSheet(ws, RangeName).Address
3         Exit Function
ErrHandler:
4         SafeRangeAddress = "#Range not found!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 19-Oct-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)

          Static InEvent As Boolean

1         On Error GoTo ErrHandler
2         If Not gBlockChangeEvent Then
3             If Not InEvent Then
4                 InEvent = True
                  'ConditionalFormatting doesn 't update as well as you might hope (because _
                   calculation is manual?) so give it a gentle kick
5                 If Not Application.Intersect(Target, RangeFromSheet(shCreditUsage, "IncludeExtraTrades")) Is Nothing Then
6                     With shCreditUsage.Range("ExtraTradeAmounts")
7                         .Value = .Value
8                     End With
9                 End If
10                If Not Application.Intersect(Target, Range(RangeFromSheet(shCreditUsage, "FilterBy1"), RangeFromSheet(shCreditUsage, "IncludeAssetClasses"))) Is Nothing Then
                      'Clear out results when the user changes which bank they are looking at...
11                    ClearoutResults
12                End If
13                If Not Application.Intersect(Target, RangeFromSheet(shCreditUsage, "ExtraTradesAre")) Is Nothing Then
14                    SetExtraTradesHeader
15                End If
16            End If
17        End If
18        InEvent = False
19        Exit Sub
ErrHandler:
20        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

