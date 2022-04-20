Attribute VB_Name = "modMarketData"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modMarketData
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Methods to grab market data, which is held on a workbook conforming to to-be-documented standards
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Function MarketBookRange(MarketBookName As String, SheetName As String, RangeName As String) As Range
          Dim R As Range
          Dim wb As Excel.Workbook
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         If Not IsInCollection(Application.Workbooks, MarketBookName) Then Throw "Workbook " + MarketBookName + " is not open"
3         Set wb = Application.Workbooks(MarketBookName)
4         If Not IsInCollection(wb.Worksheets, SheetName) Then Throw "Cannot find sheet " + SheetName + " in workbook " + MarketBookName
5         Set ws = wb.Worksheets(SheetName)
6         On Error Resume Next
7         Set R = ws.Range(RangeName)
8         On Error GoTo ErrHandler
9         If R Is Nothing Then Throw "Cannot find range named " + RangeName + " on sheet " + SheetName + " of book " + MarketBookName + "!"
10        Set MarketBookRange = R
11        Exit Function
ErrHandler:
12        Throw "#MarketBookRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketAnchorDate
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Returns the "Today Date" of the market data held in the MarketBook
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketAnchorDate(MarketBookName As String)
1         On Error GoTo ErrHandler
2         sMarketAnchorDate = MarketBookRange(MarketBookName, "DiscountFactors", "AnchorDate").Value2
3         Exit Function
ErrHandler:
4         sMarketAnchorDate = "#sMarketAnchorDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketDiscountFactor
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Interpolate discount factors (quite naive linear in zero coupon rate) off tables
'             held on sheet DiscountFactors or market book
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketDiscountFactor(Ccy As String, TheDates As Variant, MarketBookName As String)
          Dim AnchorDate
          Dim DFGridDates
          Dim DFZeros
          Dim ZeroRates

1         On Error GoTo ErrHandler
2         With MarketBookRange(MarketBookName, "DiscountFactors", "DF_" & Ccy)
3             DFGridDates = .Columns(1).Value2
4             DFZeros = .Columns(3).Value2
5         End With
6         AnchorDate = DFGridDates(1, 1)
7         ZeroRates = ThrowIfError(sInterp(DFGridDates, DFZeros, TheDates, , "NF"))
8         sMarketDiscountFactor = sArrayExp(sArrayMultiply(ZeroRates, -1, sArrayDivide((sArraySubtract(TheDates, AnchorDate)), 365)))

9         Exit Function
ErrHandler:
10        sMarketDiscountFactor = "#sMarketDiscountFactor (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketFxVol
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Interpolate FX vols from data held on the FxVol sheet of the MarketBook. FxVol USD/EUR is the
'             same as FxVol EUR/USD so we search for either range on the market sheet.
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketFxVol(Ccy1 As String, Ccy2 As String, TheDates As Variant, MarketBookName As String, Optional UseHistorical As Boolean = False, Optional WithShocks As Boolean = False)
          Dim MatchID
          Dim RangeName As String
          Dim xArrayAscending
          Dim yArray

1         On Error GoTo ErrHandler

2         If WithShocks Then
3             If UseHistorical Then
4                 RangeName = "FxVolsHistorical"
5             Else
6                 RangeName = "FxVols"
7             End If
8         Else
9             If UseHistorical Then
10                RangeName = "FxVolsHistoricalUnshocked"
11            Else
12                RangeName = "FxVolsUnShocked"
13            End If
14        End If

15        With MarketBookRange(MarketBookName, "FxVols", RangeName)
16            MatchID = sMatch(Ccy1 + Ccy2, .Columns(1).Value2)
17            If VarType(MatchID) = vbString Then
18                MatchID = sMatch(Ccy2 + Ccy1, .Columns(1).Value2)
19            End If
20            If VarType(MatchID) = vbString Then
21                Throw "Cannot find FxVolData for currency pair " + CStr(Ccy1) + CStr(Ccy2)
22            End If
23            xArrayAscending = sArrayTranspose(.Rows(1).Offset(, 1).Resize(, .Columns.Count - 1).Value2)
24            yArray = sArrayTranspose(.Rows(MatchID).Offset(, 1).Resize(, .Columns.Count - 1).Value2)
25        End With

26        sMarketFxVol = sInterp(xArrayAscending, yArray, TheDates, "Linear", "FF")

27        Exit Function
ErrHandler:
28        sMarketFxVol = "#sMarketFxVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketFxPerBaseCcy
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Return the spot Fx rate against the base currency (quoted as Ccy/BaseCcy)
'             from the data held on the DiscountFactors sheet of the MarketBook.
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketFxPerBaseCcy(ByVal TheCCys As Variant, BaseCCY As String, MarketBookName As String, Optional WithShocks As Boolean = False)
          Dim BaseCcyRate
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim MatchRes
          Dim MatchRes2
          Dim N As Long
          Dim RangeName As String
          Dim Result()
          Dim Rng As Range
          Dim SpotLevelsArray
1         On Error GoTo ErrHandler
2         Force2DArrayR TheCCys
3         N = sNRows(TheCCys)
4         M = sNCols(TheCCys)

5         If WithShocks Then
6             RangeName = "FxSpotLevels"
7         Else
8             RangeName = "FxSpotLevelsUnShocked"
9         End If

10        Set Rng = MarketBookRange(MarketBookName, "DiscountFactors", RangeName)

11        SpotLevelsArray = Rng
12        MatchRes2 = sMatch(BaseCCY, Rng.Columns(1).Value2)
13        If Not IsNumber(MatchRes2) Then Throw BaseCCY + " is not listed in range " + RangeName + " of sheet DiscountFactors of book " + MarketBookName
14        BaseCcyRate = SpotLevelsArray(MatchRes2, 3)

15        MatchRes = sMatch(TheCCys, Rng.Columns(1).Value2)
16        Force2DArray MatchRes
17        ReDim Result(1 To N, 1 To M)

18        For i = 1 To N
19            For j = 1 To M
20                If Not IsNumber(MatchRes(i, j)) Then Throw "Cannot find spot rate " + CStr(TheCCys(i, j)) + BaseCCY
21                Result(i, j) = SpotLevelsArray(MatchRes(i, j), 3) / BaseCcyRate
22            Next j
23        Next i

24        sMarketFxPerBaseCcy = Result

25        Exit Function
ErrHandler:
26        sMarketFxPerBaseCcy = "#sMarketFxPerBaseCcy (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketFxForwardRates
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Returns an array of Fx forward rates. In this version TheDates is an array
'             but Ccy and BaseCCy are strings
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketFxForwardRates(TheDates As Variant, Ccy As String, BaseCCY As String, MarketBookName As String, Optional WithShocks As Boolean = False)
          Dim Result
1         On Error GoTo ErrHandler
2         Result = sArrayMultiply(ThrowIfError(sMarketFxPerBaseCcy(Ccy, BaseCCY, MarketBookName, WithShocks)), sArrayDivide(ThrowIfError(sMarketDiscountFactor(Ccy, TheDates, MarketBookName)), ThrowIfError(sMarketDiscountFactor(BaseCCY, TheDates, MarketBookName))))
3         sMarketFxForwardRates = Result
4         Exit Function
ErrHandler:
5         sMarketFxForwardRates = "#sMarketFxForwardRates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketFxForwardRates2
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Returns an array of Fx forward rates. In this version TheDates, Ccys and
'             BaseCCys must all be arrays of the same size. COULD BE RE-WRITTEN FOR BETTER SPEED
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketFxForwardRates2(TheDates, CCys, BaseCcys, MarketBookName As String, Optional WithShocks As Boolean = False)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Force2DArrayRMulti TheDates, CCys, BaseCcys
3         NR = sNRows(TheDates): NC = sNCols(TheDates)
4         Result = sReshape(0, NR, NC)
5         For i = 1 To NR
6             For j = 1 To NC
7                 Result(i, j) = sMarketFxPerBaseCcy(CStr(CCys(i, j)), CStr(BaseCcys(i, j)), MarketBookName, WithShocks)(1, 1)
8                 Result(i, j) = Result(i, j) * sMarketDiscountFactor(CStr(CCys(i, j)), TheDates(i, j), MarketBookName)(1, 1) / sMarketDiscountFactor(CStr(BaseCcys(i, j)), TheDates(i, j), MarketBookName)(1, 1)
9             Next j
10        Next i
11        sMarketFxForwardRates2 = Result
12        Exit Function
ErrHandler:
13        sMarketFxForwardRates2 = "#sMarketFxForwardRates2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMarketCorrelationMatrix
' Author    : Philip Swannell
' Date      : 6-July-2015
' Purpose   : Interpolate a small correlation matrix out of the (heroically large)
'             correlation matrix held on the FxVols sheet of the MarketBook
' -----------------------------------------------------------------------------------------------------------------------
Function sMarketCorrelationMatrix(CCyList As Variant, BaseCCY As String, MarketBookName As String)
          Dim BigCorrMatrix
          Dim Headers
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim N As Long
          Dim Result()
1         Force2DArrayR CCyList

2         On Error GoTo ErrHandler
3         N = sNRows(CCyList)

4         With MarketBookRange(MarketBookName, "FxVols", "FxCorrelationBase" & BaseCCY)
5             Headers = .Offset(1).Resize(.Rows.Count - 1, 1).Value2
6             BigCorrMatrix = .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1).Value2
7         End With

8         MatchIDs = sMatch(CCyList, Headers)
9         Force2DArray MatchIDs
10        ReDim Result(1 To N, 1 To N)

11        For i = 1 To N
12            If Not IsNumber(MatchIDs(i, 1)) Then Throw "Cannot find currency " + CStr(CCyList(i, 1)) + " in headers of range FxCorrelationBase" & UCase$(BaseCCY) + " on sheet FxVols of book " + MarketBookName
13            For j = 1 To N
14                If Not IsNumber(MatchIDs(j, 1)) Then Throw "Cannot find currency " + CStr(CCyList(j, 1)) + " in headers of range FxCorrelationBase" & UCase$(BaseCCY) + " on sheet FxVols of book " + MarketBookName
15                Result(i, j) = BigCorrMatrix(MatchIDs(i, 1), MatchIDs(j, 1))
16            Next j
17        Next i
18        sMarketCorrelationMatrix = Result
19        Exit Function
ErrHandler:
20        sMarketCorrelationMatrix = "#sMarketCorrelationMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
