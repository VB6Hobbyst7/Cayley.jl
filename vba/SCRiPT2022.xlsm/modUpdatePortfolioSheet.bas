Attribute VB_Name = "modUpdatePortfolioSheet"
'---------------------------------------------------------------------------------------
' Module    : modUpdateSheets
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Code to get data back from the Julia environment and paste it to the Portfolio sheet
'---------------------------------------------------------------------------------------

Option Explicit

Sub TestUpdatePortfolioSheet()
1         On Error GoTo ErrHandler
2         UpdatePortfolioSheet
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestUpdatePortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdatePortfolioSheet
' Author    : Philip Swannell
' Date      : 13-Apr-2016
' Purpose   : Faster way (versus method XVAFrontEndMain) of updating the Portfolio sheet
'             Calls R method DataForPortfolioSheet and only works if the model already exists in R - see method ModelExists
'---------------------------------------------------------------------------------------
Sub UpdatePortfolioSheet()
1         On Error GoTo ErrHandler

          Dim c As Range
          Dim CallToRFailed As Boolean
          Dim CopyOfErr As String
          Dim ErrorString As String
          Dim NumTrades As Long
          Dim OldBCE As Boolean
          Dim ReportCurrencies
          Dim ResultsFromJulia As Variant
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim Trades As Variant
          Dim TradesRange As Range

2         OldBCE = gBlockChangeEvent
3         gBlockChangeEvent = True

4         Set TradesRange = getTradesRange(NumTrades)
5         If NumTrades = 0 Then GoTo EarlyExit

6         If TradeIDsNeedRepairing(True) Then Throw "Duplicate TradeIDs exist. Fix with" + vbLf + " Menu > Trades > Repair invalid Trade IDs", True

7         CalculatePortfolioSheet
8         Trades = TradesRange.Value2
9         ReportCurrencies = TradesRange.Columns(gCN_Ccy1).Value
          If TradesRange.Rows.Count = 1 Then Force2DArray ReportCurrencies

10        ResultsFromJulia = DataForPortfolioSheetVBA(ReportCurrencies)

11        If VarType(ResultsFromJulia) = vbString Then
12            CallToRFailed = True
13        Else
              Dim DataToPaste
14            DataToPaste = ResultsFromJulia
15            Force2DArray DataToPaste
16        End If

17        If CallToRFailed Then
18            ErrorString = CStr(ResultsFromJulia)
19        End If

20        Set SUH = CreateScreenUpdateHandler()
21        Set SPH = CreateSheetProtectionHandler(shPortfolio)

22        With TradesRange

23            If CallToRFailed Then
24                .Columns(.Columns.Count - 2).Resize(, 2).Value = ""
25                .Columns(.Columns.Count).Value = "'" + ErrorString    'does left alignment
26            Else
27                .Columns(.Columns.Count - 2).Resize(, 3).Value = DataToPaste
28                .Columns(.Columns.Count).HorizontalAlignment = xlHAlignCenter
29                For Each c In .Columns(.Columns.Count).Cells
30                    If VarType(c.Value) = vbString Then
31                        c.HorizontalAlignment = xlHAlignLeft
32                    End If
33                Next c
34            End If

35        End With
36        RangeFromSheet(shPortfolio, "TotalPV").Calculate
37        SetTradesRangeColumnWidths
38        FilterTradesRange

EarlyExit:
39        gBlockChangeEvent = OldBCE
40        Exit Sub
ErrHandler:
41        CopyOfErr = "#UpdatePortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
42        gBlockChangeEvent = OldBCE
43        Throw CopyOfErr
End Sub

Sub TestDataForPortfolioSheetVBA()
1         On Error GoTo ErrHandler
2         If gResults Is Nothing Then ReloadResults

5         DataForPortfolioSheetVBA "Foo"
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#TestDataForPortfolioSheetVBA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'accesses global gResults. TODO - pass down call stack?
Function DataForPortfolioSheetVBA(ReportCurrencies)
          Dim FxRates As Dictionary
          Dim i As Long
          Dim N As Long
          Dim PVs
          Dim Result
          Dim Statuses
          
1         On Error GoTo ErrHandler
2         PVs = gResults("TradeResults")("PV")
3         Statuses = gResults("TradeResults")("PVStatus")
4         N = UBound(PVs)
5         Result = sReshape(0, N, 3)
          
6         Set FxRates = gResults("Model")("spot")

7         For i = 1 To N
              'Note change of Fx quote convention
8             Result(i, 1) = 1 / FxRates(ReportCurrencies(i, 1))
9             If Statuses(i) = "OK" Then
                  'Note change of sign
10                Result(i, 2) = -PVs(i) / FxRates(ReportCurrencies(i, 1))
11                Result(i, 3) = -PVs(i)
12            Else
13                Result(i, 2) = Statuses(i)
14                Result(i, 3) = Statuses(i)
15            End If
16        Next i

17        DataForPortfolioSheetVBA = Result

18        Exit Function
ErrHandler:
19        Throw "#DataForPortfolioSheetVBA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'DataForPortfolioSheet = function(Model, Trades = NULL, TradeResults = NULL, ReportCurrencies) {
'    out <- tryCatch({
'        if (is.null(TradeResults)) {
'            TradeResults = CreateTradeResults(Trades)
'            TradeResults = DoTaskTradePV(Model, Trades, TradeResults, TRUE, FALSE)
'            TradeResults[, c("DVA", "CVA", "PVWhatIf", "DVAWhatIf", "CVAWhatIf", "FundingPVWhatIf")] = 0
'            gCachedDataForDashboardInPVOnlyMode <<- cbind(Trades[c("TradeID", "ValuationFunction", "Counterparty")], TradeResults[, c("PV", "FundingPV")])
'        }
'        ThirdCol = -TradeResults[, "PV"] #Note sign flip, since Portfolio sheet displays PVs from banks' perspective
'        FxRates = unlist(Model$spot[Model$Currencies])
'        names(FxRates) = Model$Currencies
'
'        #When reporting value of inflation trades, quote in the base currency of that index...
'        if (length(Model$Inflations) > 0) {
'            for (i in 1:length(Model$Inflations)) {
'                base = Model$parameters[[Model$Inflations[[i]]]]$BaseCurrency
'                rate = Model$spot[[base]]
'                FxRates[[Model$Inflations[[i]]]] = rate
'            }
'        }
'        FirstCol = 1 / FxRates[ReportCurrencies]
'        SecondCol = ThirdCol * FirstCol
'        Result = cbind(FirstCol, SecondCol, ThirdCol)
'        if (!all(TradeResults$PVStatus == "OK")) {
'            ResultWithErrors = matrix(list(), nrow(TradeResults), 3)
'            for (i in 1:nrow(TradeResults)) {
'                if (TradeResults[i, "PVStatus"] == "OK") {
'                    ResultWithErrors[i, 1] = list(Result[i, 1])
'                    ResultWithErrors[i, 2] = list(Result[i, 2])
'                    ResultWithErrors[i, 3] = list(Result[i, 3])
'                }
'                else {
'                    ResultWithErrors[i, 1] = list(Result[i, 1])
'                    ResultWithErrors[i, 2] = list(TradeResults[i, "PVStatus"])
'                    ResultWithErrors[i, 3] = list(TradeResults[i, "PVStatus"])
'                }
'            }
'            Result <- ResultWithErrors
'        }
'        dimnames(Result) <- NULL #otherwise when passed back to Excel via BERT, the dimnames appear as headers
'        Result
'    }, error = function(e) {
'        stop(AddContext("DataForPortfolioSheet", e))
'    },
'  warning = function(e) {
'    stop(AddContext("DataForPortfolioSheet", e, "Warning"))
'  })
'    return(out)
'}

'---------------------------------------------------------------------------------------
' Procedure : MultiAreaValue
' Author    : Philip Swannell
' Date      : 28-Apr-2016
' Purpose   : Taking the .Value2 property of a mult-area range yields only the .Value2 of the
'             first area. This method returns an array-stack of the values of the areas.
'---------------------------------------------------------------------------------------
Function MultiAreaValue2(R As Range)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res() As Variant
          Dim TCNC As Long
          Dim TCNR As Long
          Dim ThisChunk
          Dim WriteOffset As Long
1         On Error GoTo ErrHandler
2         If R.Areas.Count = 1 Then
3             MultiAreaValue2 = R.Value2
4             Exit Function
5         Else
6             NR = R.Areas(1).Rows.Count
7             NC = R.Areas(1).Columns.Count
8             For k = 2 To R.Areas.Count
9                 NR = NR + R.Areas(k).Rows.Count
10                If NC < R.Areas(k).Columns.Count Then NC = R.Areas(k).Columns.Count
11            Next k
12            ReDim Res(1 To NR, 1 To NC)

13            WriteOffset = 0
14            For k = 1 To R.Areas.Count
15                ThisChunk = R.Areas(k).Value2
16                TCNR = sNRows(ThisChunk)
17                TCNC = sNCols(ThisChunk)
18                For i = 1 To TCNR
19                    For j = 1 To TCNC
20                        Res(i + WriteOffset, j) = ThisChunk(i, j)
21                    Next j
22                Next i
23                WriteOffset = WriteOffset + TCNR
24            Next k
25        End If
26        MultiAreaValue2 = Res
27        Exit Function
ErrHandler:
28        Throw "#MultiAreaValue2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
