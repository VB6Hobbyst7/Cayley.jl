Attribute VB_Name = "modSolve"
'---------------------------------------------------------------------------------------
' Module    : modSolve
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Methods for solving trades for zero PV, available from the right-click menu
'             on the Portfolio sheet.
'---------------------------------------------------------------------------------------
Option Explicit

Function SolveSelectedTradesForTargetPV()
          Static TargetPV As Double
          Static TargetCurrency As String
          Const Title = "Solve trades"
          Dim ButtonClicked As String
          Dim Res1 As Variant

1         On Error GoTo ErrHandler
EnterCCy:
2         Res1 = ShowOptionButtonDialog(sSortedArray(CurrenciesInModel()), Title, "Select currency of target PV", TargetCurrency, , , , , , "Next >", , , ButtonClicked)

3         If ButtonClicked = "Cancel" Then
4             Exit Function
5         Else
6             TargetCurrency = Res1
7         End If

EnterAmount:
8         Res1 = InputBoxPlus("Enter target for PV (from Bank PoV) in " & TargetCurrency, Title, CStr(TargetPV), "< Back", , , , , , , "OK", ButtonClicked)

9         If ButtonClicked = "Cancel" Then
10            Exit Function
11        ElseIf ButtonClicked = "< Back" Then
12            GoTo EnterCCy
13        Else
14            If IsNumeric(Res1) Then
15                TargetPV = CDbl(Res1)
16            Else
17                GoTo EnterAmount
18            End If
19        End If

20        SolveSelectedTrades -TargetPV, TargetCurrency

21        Exit Function
ErrHandler:
22        Throw "#SolveSelectedTradesForTargetPV (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

'---------------------------------------------------------------------------------------
' Procedure : SolveSelectedTrades
' Author    : Philip Swannell
' Date      : 12-Apr-2016
' Purpose   : For each cell currently selected change the trade attribute at that cell
'             so as to make the trade have zero value, callable from the right-click menu.
'---------------------------------------------------------------------------------------
Function SolveSelectedTrades(Optional TargetPV As Double = 0, Optional TargetCurrency As String)
          Dim c As Range
          Dim CellsToProcess As Range
          Dim ColNoTradeID
          Dim i As Long
          Dim NumFailed As Long
          Dim NumTrades As Long
          Dim NumWorked As Long
          Dim Prompt As String
          Dim PromptFail As String
          Dim PromptSuccess As String
          Dim Res As Variant
          Dim ThisHeader As String
          Dim ThisTrade As Variant
          Dim TradesRange As Range
          Dim SuccessFlags As Variant    ' SuccessFlags avoid solving the same trade twice
          Dim CopyOfErr As String
          Dim OldBCE As Boolean

1         On Error GoTo ErrHandler
2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         ColNoTradeID = gCN_TradeID

5         Set TradesRange = getTradesRange(NumTrades)
6         SuccessFlags = sReshape(False, NumTrades, 1)

7         On Error Resume Next
8         Set CellsToProcess = Application.Intersect(TradesRange, Selection)
9         On Error GoTo ErrHandler

10        Set CellsToProcess = UnhiddenRowsInRange(CellsToProcess)

11        For Each c In CellsToProcess.Cells
12            If Not SuccessFlags(c.Row - TradesRange.Row + 1, 1) Then
13                ThisTrade = TradesRange.Rows(c.Row - TradesRange.Row + 1).Value2
                  'Replace empties with zeros makes it possible to solve for a field that is currently empty
14                For i = 1 To sNCols(ThisTrade)
15                    If IsEmpty(ThisTrade(1, i)) Then ThisTrade(1, i) = 0#
16                Next i

17                ThisHeader = TradesRange.Cells(0, c.Column - TradesRange.Column + 1).Value2
18                Res = TradeSolver(ThisTrade, ThisHeader, TargetPV, False, TargetCurrency)
19                If VarType(Res) <> vbString Then
20                    SuccessFlags(c.Row - TradesRange.Row + 1, 1) = True
21                    NumWorked = NumWorked + 1
22                    PromptSuccess = PromptSuccess + vbLf + "TradeID '" + CStr(ThisTrade(1, ColNoTradeID)) + "' solved: " + ThisHeader + " = " + CStr(Res)
23                    c.Value = Res
24                Else
25                    NumFailed = NumFailed + 1
26                    PromptFail = PromptFail + vbLf + "TradeID '" + CStr(ThisTrade(1, ColNoTradeID)) + "' failed: " + Res
27                End If
28            End If
29        Next

30        If NumWorked > 0 Then
31            UpdatePortfolioSheet
32        End If

33        If NumFailed = 0 Then
34            Prompt = CStr(NumWorked) + " trade" + IIf(NumWorked = 1, "", "s") + " solved for " + IIf(TargetPV = 0, "zero", "target") + " PV."
35            TemporaryMessage Prompt, 5
36        Else
              Dim TextWidth As Double
37            Prompt = CStr(NumWorked) + " trade" + IIf(NumWorked = 1, "", "s") + " solved for " + IIf(TargetPV = 0, "zero", "target") + " PV" + _
                  IIf(NumFailed = 0, ".", ", " + CStr(NumFailed) + " failure" + IIf(NumFailed = 1, "", "s") + ".")
38            If NumWorked > 0 Then
39                Prompt = Prompt + vbLf + vbLf + "Solved trades:" + PromptSuccess
40            End If
41            Prompt = Prompt + vbLf + vbLf + "Failures:" + PromptFail
42            TextWidth = sColumnMax(sStringWidth(sTokeniseString(Prompt, vbLf), "Calibri", 11))(1, 1) + 20
43            If TextWidth > 800 Then TextWidth = 800
44            MsgBoxPlus Prompt, , MsgBoxTitle() + " Trade Solver", , , , , CLng(TextWidth)
45        End If

EarlyExit:
46        gBlockChangeEvent = OldBCE

47        Exit Function
ErrHandler:
48        CopyOfErr = "#SolveSelectedTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
49        gBlockChangeEvent = OldBCE
50        SomethingWentWrong CopyOfErr
End Function

'---------------------------------------------------------------------------------------
' Procedure : AreSelectedTradesSolvable
' Author    : Philip Swannell
' Date      : 28-Apr-2016
' Purpose   : Figure out if at least one of the selected trades is solvable
'---------------------------------------------------------------------------------------
Function AreSelectedTradesSolvable()
          Dim CellsToProcess As Range
          Dim NumTrades As Long
          Dim Res As Variant
          Dim ThisHeader As String
          Dim ThisTrade As Variant
          Dim TradesRange As Range

          Dim ColNoTradeID
          Dim ColNoValuationFunction

1         On Error GoTo ErrHandler

2         If Not ModelExists() Then
3             AreSelectedTradesSolvable = False
4             Exit Function
5         End If

6         ColNoTradeID = gCN_TradeID
7         ColNoValuationFunction = gCN_TradeType

8         Set TradesRange = getTradesRange(NumTrades)
9         If NumTrades = 0 Then
10            AreSelectedTradesSolvable = False
11            Exit Function
12        End If

13        On Error Resume Next
14        Set CellsToProcess = Application.Intersect(TradesRange, Selection)
15        On Error GoTo ErrHandler

16        If CellsToProcess Is Nothing Then
17            AreSelectedTradesSolvable = False
18            Exit Function
19        End If

20        If Application.Intersect(CellsToProcess.EntireColumn, shPortfolio.Rows(1)).Cells.Count > 1 Then
21            AreSelectedTradesSolvable = False    'This restriction makes it much easier for this method to be fast
22            Exit Function
23        End If

24        Set CellsToProcess = UnhiddenRowsInRange(CellsToProcess)

25        If CellsToProcess Is Nothing Then
26            AreSelectedTradesSolvable = False
27            Exit Function
28        End If

          Dim AllTrades
          Dim ChooseVector
          Dim i As Long
          Dim VFs

          'Since we are resticting to the case when the selection is a single column, we only need to test the first example of a each type of trade (each VF)
29        AllTrades = MultiAreaValue2(Application.Intersect(CellsToProcess.EntireRow, TradesRange))
30        VFs = sSubArray(AllTrades, 1, ColNoValuationFunction, , 1)
31        ChooseVector = sArrayEquals(sMatch(VFs, VFs), sIntegers(sNRows(VFs)))
32        ThisHeader = TradesRange.Cells(0, CellsToProcess.Column - TradesRange.Column + 1).Value2
33        For i = 1 To sNRows(AllTrades)
34            If ChooseVector(i, 1) Then
35                ThisTrade = sSubArray(AllTrades, 1, 1, 1)

36                Res = TradeSolver(ThisTrade, ThisHeader, 0, True)

37                If VarType(Res) = vbBoolean Then
38                    If Res Then
39                        AreSelectedTradesSolvable = True
40                        Exit Function
41                    End If
42                End If
43            End If
44        Next

45        AreSelectedTradesSolvable = False
46        Exit Function
ErrHandler:
47        Throw "#AreSelectedTradesSolvable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : TradeSolver
' Author    : Philip Swannell
' Date      : 12-Apr-2016
' Purpose   : Wrap to the R function TradeSolver
'---------------------------------------------------------------------------------------
Function TradeSolver(Trade, Header, Optional TargetPV As Double = 0, Optional JustTestForSolvability As Boolean, Optional TargetCurrency As String)
          Dim ChangeSign As Boolean
          Dim EvaluateThis As String
          Dim Res As Variant
          Dim SearchFrom As Double
          Dim SearchTo As Double
          Dim TradeAttribute As String
          Dim TradeForR
          Dim VF As String

1         On Error GoTo ErrHandler
2         VF = Trade(1, gCN_TradeType)

3         Select Case VF & "|" & Header
              Case "InterestRateSwap|Rate 1"
4                 If Trade(1, gCN_LegType1) = "Fixed" Then
5                     TradeAttribute = "Coupon"
6                     SearchFrom = -1: SearchTo = 1
7                 Else
8                     TradeAttribute = "Margin"
9                     SearchFrom = -1: SearchTo = 1
10                End If
11            Case "InterestRateSwap|Rate 2"
12                If Trade(1, gCN_LegType2) = "Fixed" Then
13                    TradeAttribute = "Coupon"
14                    SearchFrom = -1: SearchTo = 1
15                Else
16                    TradeAttribute = "Margin"
17                    SearchFrom = -1: SearchTo = 1
18                End If
19            Case "InflationZCSwap|Rate 1"
20                If Trade(1, gCN_LegType1) = "Fixed" Then
21                    TradeAttribute = "Coupon"
22                    SearchFrom = -1: SearchTo = 1
23                Else
24                    Throw "Not solvable"
25                End If
26            Case "InflationZCSwap|Rate 2"
27                If Trade(1, gCN_LegType2) = "Fixed" Then
28                    TradeAttribute = "Coupon"
29                    SearchFrom = -1: SearchTo = 1
30                Else
31                    Throw "Not solvable"
32                End If
33            Case "InflationYoYSwap|Rate 1"
34                Select Case Trade(1, gCN_LegType1)
                      Case "Fixed", "Floating"
35                        TradeAttribute = "ReceiveCoupon"
36                        SearchFrom = -1: SearchTo = 1
37                    Case Else
38                        Throw "Not solvable"
39                End Select
40            Case "InflationYoYSwap|Rate 2"
41                Select Case Trade(1, gCN_LegType2)
                      Case "Fixed", "Floating"
42                        TradeAttribute = "PayCoupon"
43                        SearchFrom = -1: SearchTo = 1
44                    Case Else
45                        Throw "Not solvable"
46                End Select
47            Case "FxForward|Notional 1"
48                TradeAttribute = "ReceiveNotional"
49                SearchFrom = 0: SearchTo = 100000000000000#
50            Case "FxForward|Notional 2"
51                TradeAttribute = "PayNotional"
52                SearchFrom = 0: SearchTo = -100000000000000#
53                ChangeSign = True    'Consequence of different sign conventions for trades held on Portfolio sheet and trades when sent to Julia
54            Case "CrossCurrencySwap|Rate 1"
55                TradeAttribute = "ReceiveCoupon"
56                SearchFrom = -100: SearchTo = 100
57            Case "CrossCurrencySwap|Rate 2"
58                TradeAttribute = "PayCoupon"
59                SearchFrom = -100: SearchTo = 100
60            Case "Swaption|Rate 1"
61                TradeAttribute = "Strike"
62                SearchFrom = 0: SearchTo = 1
63            Case "CapFloor|Rate 1"
64                TradeAttribute = "Strike"
65                SearchFrom = 0: SearchTo = 1
66            Case Else
67                Throw "Solving not implemented for TradeType = " & VF & ", Column Header = " & Header
68        End Select

69        If JustTestForSolvability Then
70            TradeSolver = True
71            Exit Function
72        End If

73        If Not ModelExists() Then
              '     XVAFrontEndMain False, False, False, False, False, False, False, False, AP_BERT
74        End If

75        TradeForR = PortfolioTradesToJuliaTrades(Trade, True, False)
76        SaveDataframe TradeForR, "input_tradetosolve"
          '      EvaluateThis = "TradeSolver(" + gModel + ",input_tradetosolve,""" + TradeAttribute + """," + CStr(TargetPV) + "," + CStr(SearchFrom) & "," + CStr(SearchTo) & ",""" + CStr(TargetCurrency) & """)"
          '     Res = ThrowIfError(sExecuteRCode(EvaluateThis))
77        TradeSolver = Res * IIf(ChangeSign, -1, 1)

78        Exit Function
ErrHandler:
79        TradeSolver = "#TradeSolver (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

