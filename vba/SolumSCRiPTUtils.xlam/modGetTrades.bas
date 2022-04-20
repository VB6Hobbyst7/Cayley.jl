Attribute VB_Name = "modGetTrades"
Option Explicit
Option Private Module

'Common sub-routines of GetTradesInRFormat and GetTradesInJuliaFormat

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetColumnsFromTradesWorkbook
' Author    : Philip Swannell
' Date      : 13-Jul-2016
' Purpose   : Gets multiple columns of trade data. Pass in header names as a column array
'             or a comma delimited string.
' -----------------------------------------------------------------------------------------------------------------------
Function GetColumnsFromTradesWorkbook(ByVal HeaderNames As Variant, IncludeFutureTrades As Boolean, _
                                      PortfolioAgeing As Double, WithFxTrades, WithRatesTrades, twb As Workbook, _
                                      fts As Worksheet, AnchorDate As Date)
          Dim i As Long
          Dim j As Long
          Dim NumCols As Long
          Dim NumRows As Long
          Dim Result() As Variant
          Dim ThisCol As Variant

1         On Error GoTo ErrHandler
2         If VarType(HeaderNames) = vbString Then HeaderNames = sTokeniseString(CStr(HeaderNames))
3         Force2DArray HeaderNames
4         NumCols = sNRows(HeaderNames)
5         NumRows = GetColumnFromTradesWorkbook("NumTrades", IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
6         ReDim Result(1 To NumRows, 1 To NumCols)

7         For j = 1 To NumCols
8             ThisCol = GetColumnFromTradesWorkbook(HeaderNames(j, 1), IncludeFutureTrades, PortfolioAgeing, WithFxTrades, WithRatesTrades, twb, fts, AnchorDate)
9             For i = 1 To NumRows
10                Result(i, j) = ThisCol(i, 1)
11            Next i
12        Next j
13        GetColumnsFromTradesWorkbook = Result
14        Exit Function
ErrHandler:
15        Throw "#GetColumnsFromTradesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DealTypeToValuationFunction
' Author    : Philip Swannell
' Date      : 03-Aug-2016
' Purpose   : Translate the strings that appear in the DEAL_TYPE column of the Airbus
'             trade data to the string to appear in the ValuationFunction column of the
'             trade data in R format.
' -----------------------------------------------------------------------------------------------------------------------
Function DealTypeToValuationFunction(DealType As String)
1         On Error GoTo ErrHandler
2         Select Case DealType
          Case "FXForward_buy", "FXForward_sell", "FXSwap_buy", "FXSwap_sell", "FXSpot_sell", "FXSpot_buy", "FxForward", "Forward"
3             DealTypeToValuationFunction = "FxForward"
4         Case "CALLbuy VANILLA", "CALLsell VANILLA"
5             DealTypeToValuationFunction = "FxOption"
6         Case "PUTbuy VANILLA", "PUTsell VANILLA"
7             DealTypeToValuationFunction = "FxOption"
8         Case "Swap"
9             DealTypeToValuationFunction = "InterestRateSwap"
10        Case "XCCySwap"
11            DealTypeToValuationFunction = "CrossCurrencySwap"
12        Case Else
13            Throw "Unrecognised value in ""DEAL_TYPE"" column: '" + DealType + "'"
14        End Select
15        Exit Function
ErrHandler:
16        Throw "#DealTypeToValuationFunction (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddTwoDays
' Author    : Philip Swannell
' Date      : 12-Dec-2016
' Purpose   : Adds two weekdays to a date:
'             Monday > Wednesday
'             Tueday > Thursday
'             Wednesday > Friday
'             Thursday > Monday
'             Friday > Tuesday
'             Saturday > Tuesday
'             Sunday > Tuesday
' -----------------------------------------------------------------------------------------------------------------------
Function AddTwoDays(inDate As Long)
1         On Error GoTo ErrHandler
2         AddTwoDays = inDate + Choose((inDate Mod 7) + 1, 3, 2, 2, 2, 2, 4, 4)
3         Exit Function
ErrHandler:
4         Throw "#AddTwoDays (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StringWithCommaToNumber
' Author    : Philip Swannell
' Date      : 02-Aug-2016
' Purpose   : The files saved by Airbus (in columns INDEX_REC, INDEX_PAY,
'             SPREAD_REC, SPREAD_PAY) represent numbers as strings with comma as the decimal
'             seperator. :-( Have asked Airbus to correct this and if they do so this method
'             will be redundant, though safe to remain in place.
' -----------------------------------------------------------------------------------------------------------------------
Function StringWithCommaToNumber(TheString As Variant)
          Static DecimalPoint As String

1         On Error GoTo ErrHandler
2         If DecimalPoint = "" Then
3             DecimalPoint = Mid(CStr(2.2), 2, 1)        ' not sure if Windows international settings can ever make this be other than "."
4         End If

5         If VarType(TheString) = vbString Then
6             StringWithCommaToNumber = CDbl(Replace(TheString, ",", DecimalPoint))
7         ElseIf IsNumber(TheString) Then
8             StringWithCommaToNumber = CDbl(TheString)
9         Else
10            Throw "Unexpected type"
11        End If

12        Exit Function
ErrHandler:
13        Throw "#StringWithCommaToNumber cannot convert '" + CStr(TheString) + "' to a number!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ParseAirbusBDC
' Author    : Philip Swannell
' Date      : 15-Sep-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function ParseAirbusBDC(BDC As String) As Variant
1         On Error GoTo ErrHandler

2         Select Case BDC
          Case "MOD_FOLLOW", "MOD_FOLLOWING"
3             ParseAirbusBDC = "Mod Foll"
4         Case "FOLLOWING"
5             ParseAirbusBDC = "Foll"
6         Case "MOD_PRECEDE", "MOD_PRECEDING"        ' Not seen any examples of this, so guessing what Airbus might use
7             ParseAirbusBDC = "Mod Prec"
8         Case "PRECEDING", "PRECEDE"        ' Not seen any examples of this
9             ParseAirbusBDC = "Prec"
10        Case Else
11            Throw "Frequency not recognised. Allowed values: MOD_FOLLOW, FOLLOWING, MOD_PRECEDE, PRECEDING"
12        End Select
13        Exit Function
ErrHandler:
14        Throw "#ParseAirbusBDC (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ParseAirbusFrequency
' Author    : Philip Swannell
' Date      : 15-Sep-2016
' Purpose   : Convert strings used for coupon frequencies in Airbus's trade data...
' -----------------------------------------------------------------------------------------------------------------------
Function ParseAirbusFrequency(FrequencyString As String) As Variant
1         On Error GoTo ErrHandler

2         Select Case FrequencyString
          Case "PA"
3             ParseAirbusFrequency = 1
4         Case "SA"
5             ParseAirbusFrequency = 2
6         Case "QTR"
7             ParseAirbusFrequency = 4
8         Case "MTH"
9             ParseAirbusFrequency = 12
10        Case Else
11            Throw "Frequency not recognised. Allowed values: PA, SA, QTR, MTH"
12        End Select
13        Exit Function
ErrHandler:
14        Throw "#ParseAirbusFrequency (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GrabAmortisationData
' Author    : Philip Swannell
' Date      : 04-Aug-2016
' Purpose   : Populates variables passed by reference. First four are simply the columns
'             in the amortisation sheet of the trades workbook and the last, AmortMap,
'             provides a map into the other four in the sense of sCountRepeats CFH on TradeIDs.
' -----------------------------------------------------------------------------------------------------------------------
Function GrabAmortisationData(ByRef AmortTradeIDs As Variant, ByRef AmortStartDates As Variant, ByRef AmortPayRecs As Variant, ByRef AmortNotionals As Variant, ByRef AmortMap As Variant, twb As Workbook)
          Dim cn_Notional
          Dim cn_PAY_REC_LEG
          Dim cn_START_DATE
          Dim cn_TRADE_ID
          Dim Headers As Variant
          Dim RngHeaders As Range
          Dim i As Long
          Dim NumRows As Long
          Dim ws As Worksheet
          Dim SheetName As String
          Dim Is2022Format As Boolean

1         On Error GoTo ErrHandler

2         If IsInCollection(twb.Worksheets, SN_Amortisation) Then
3             SheetName = SN_Amortisation
4             Is2022Format = False
5         ElseIf IsInCollection(twb.Worksheets, SN_Amortisation2) Then
6             Is2022Format = True
7             SheetName = SN_Amortisation2
8         Else
9             Throw "Cannot find sheet containing amortisation data in workbook " + twb.Name + " we looked for sheets named either '" & SN_Amortisation & "' or '" & SN_Amortisation2 & "'"
10        End If
11        Set ws = twb.Worksheets(SheetName)
12        If Is2022Format Then
13            Set RngHeaders = ws.ListObjects(1).DataBodyRange.Rows(0)
14        Else
15            Set RngHeaders = sExpandRight(ws.Cells(1, 1))
16        End If

17        Headers = sArrayTranspose(RngHeaders.Value)

18        cn_TRADE_ID = sMatch("TRADE_ID", Headers): If Not IsNumber(cn_TRADE_ID) Then Throw "Cannot find 'TRADE_ID' in top row of worksheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "'"
19        cn_START_DATE = sMatch("START_DATE", Headers): If Not IsNumber(cn_START_DATE) Then Throw "Cannot find 'START_DATE' in top row of worksheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "'"
20        cn_PAY_REC_LEG = sMatch("PAY_REC_LEG", Headers): If Not IsNumber(cn_PAY_REC_LEG) Then Throw "Cannot find 'PAY_REC_LEG' in top row of worksheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "'"
21        cn_Notional = sMatch("NOTIONAL", Headers): If Not IsNumber(cn_Notional) Then Throw "Cannot find 'NOTIONAL' in top row of worksheet '" + ws.Name + "' in workbook '" + ws.Parent.Name + "'"

22        If Is2022Format Then
23            NumRows = ws.ListObjects(1).DataBodyRange.Rows.Count
24        Else
25            NumRows = sExpandDown(Headers.Cells(1, cn_TRADE_ID)).Rows.Count - 1
26        End If

27        If NumRows = 0 Then
28            Exit Function
29        End If

30        AmortTradeIDs = RngHeaders.Cells(2, cn_TRADE_ID).Resize(NumRows)
31        AmortStartDates = RngHeaders.Cells(2, cn_START_DATE).Resize(NumRows)
32        AmortPayRecs = RngHeaders.Cells(2, cn_PAY_REC_LEG).Resize(NumRows)
33        AmortNotionals = RngHeaders.Cells(2, cn_Notional).Resize(NumRows)
34        AmortMap = sCountRepeats(AmortTradeIDs, "CFH")
35        For i = 2 To NumRows
36            If AmortTradeIDs(i, 1) < AmortTradeIDs(i - 1, 1) Then Throw "Data in worksheet '" + SN_Amortisation + "' of workbook '" + twb.Name + "' must be sorted in ascending order by TRADE_ID column"
37        Next

38        Exit Function
ErrHandler:
39        Throw "#GrabAmortisationData (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InferNotionalSchedule
' Author    : Philip Swannell
' Date      : 20-Sep-2016
' Purpose   : Data we receive from Airbus has for each amortising trade a list of dates
'             and a list of notionals, but the dates may only be a subset of the swap
'             payment dates (e.g. if the trade notional stays constant for a number of
'             periods and then declines). This function does the necessary interpolation etc
'             to get a one-notional-per coupon list.
'             The EndDate passed in should be have been adjusted by PortfolioAgeing * 365 and therefore may not be a whole number
' -----------------------------------------------------------------------------------------------------------------------
Function InferNotionalSchedule(ByVal StartDate, ByVal EndDate, ByVal Frequency As Variant, BDC As String, ByVal NotionalDates, ByVal NotionalAmounts, PortfolioAgeing As Double)
          Dim AgedSwapStartDates As Variant
          Dim AnyExcluded As Boolean
          Dim ChooseVector As Variant
          Dim i As Long
          Dim NR As Long
          Dim originalEndDate As Long
          Dim originalSwapStartDates
          Dim Result

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti NotionalDates, NotionalAmounts

9         If CLng(EndDate) <= StartDate Then
10            InferNotionalSchedule = sTake(NotionalAmounts, -1)
11            Exit Function
12        End If

13        originalEndDate = CLng(EndDate + PortfolioAgeing * 365)

14        originalSwapStartDates = ThrowIfError(DateSchedule(CDate(StartDate), originalEndDate, Frequency, BDC, "StartDates"))
15        NotionalDates = FindClosest(NotionalDates, originalSwapStartDates)        'because Airbus's dates and our dates may differ (e.g. bank holidays)
          'Doing FindClosest can snap two of the NotionalDates to the same date
16        NR = sNRows(NotionalDates)
17        ChooseVector = sReshape(True, NR, 1)
18        For i = 2 To NR
19            If NotionalDates(i, 1) <= NotionalDates(i - 1, 1) Then
20                ChooseVector(i, 1) = False
21                AnyExcluded = True
22            End If
23        Next i
24        If AnyExcluded Then
25            NotionalDates = sMChoose(NotionalDates, ChooseVector)
26            NotionalAmounts = sMChoose(NotionalAmounts, ChooseVector)
27        End If

28        Result = ThrowIfError(sInterp(NotionalDates, NotionalAmounts, originalSwapStartDates, "FlatFromLeft", "FF"))

29        If PortfolioAgeing > 0 Then
30            AgedSwapStartDates = ThrowIfError(DateSchedule(CDate(StartDate), CLng(EndDate), Frequency, BDC))
31            If sNRows(AgedSwapStartDates) < sNRows(originalSwapStartDates) Then
32                Result = sTake(Result, -sNRows(AgedSwapStartDates))
33            End If
34        End If
35        InferNotionalSchedule = Result

36        Exit Function
ErrHandler:
37        Throw "#InferNotionalSchedule (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FindClosest
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : TargetValues and AllowedValues should be column arrays of numbers and
'             AllowedValues must be sorted. Return has same number of rows as TargetValues
'             and each element of the return is that element of AllowedValues which is closest
'             to the corresponding element in TargetValues.
' -----------------------------------------------------------------------------------------------------------------------
Function FindClosest(TargetValues, AllowedValues)
          Dim i As Long
          Dim InterpRes
          Dim nAllowed As Long
          Dim nTargets As Long
          Dim Res
          Dim x As Double
          Dim y As Double

1         On Error GoTo ErrHandler
2         nAllowed = sNRows(AllowedValues)
3         nTargets = sNRows(TargetValues)
4         Res = sReshape(0, nTargets, 1)

5         InterpRes = sInterp(AllowedValues, sIntegers(nAllowed), TargetValues, "FlatFromLeft", "FF")

6         For i = 1 To nTargets
7             If InterpRes(i, 1) = nAllowed Then
8                 Res(i, 1) = AllowedValues(nAllowed, 1)
9             Else
10                x = AllowedValues(InterpRes(i, 1), 1)
11                y = AllowedValues(InterpRes(i, 1) + 1, 1)
12                If Abs(TargetValues(i, 1) - x) < Abs(TargetValues(i, 1) - y) Then
13                    Res(i, 1) = x
14                Else
15                    Res(i, 1) = y
16                End If
17            End If
18        Next i
19        FindClosest = Res
20        Exit Function
ErrHandler:
21        Throw "#FindClosest (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

