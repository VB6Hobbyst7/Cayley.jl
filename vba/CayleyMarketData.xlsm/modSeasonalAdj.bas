Attribute VB_Name = "modSeasonalAdj"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : EstimateSeasonalAdjustments
' Author    : Philip Swannell
' Date      : 25-Apr-2017
' Purpose   : Philip's patented method from estimating seasonal adjustments from historic time series.
' Problem 1:  Presumably a lot less sophisticated than method than X-12-ARIMA
'             https://www.ons.gov.uk/ons/guide-method/method-quality/general-methodology/time-series-analysis/guide-to-seasonal-adjustment.pdf
' Problem 2: Banks might like to infer seasonal adjustment from market prices for broken-date swaps, not infer from historic data
' These seasonal adjustments are as per "Multiplicative case" in a paper from OpenGamma - "Inflation: Instruments and cuve construction"
' Sum of the 12 elements in the return must be 0
' See also companion function RemoveSeasonality
' -----------------------------------------------------------------------------------------------------------------------
Function EstimateSeasonalAdjustments(ByVal InputData As Variant, Optional UseThisManyMonths As Long, Optional ByRef DataUsed, Optional ByRef FirstMonthUsed)
          Dim i As Long
          Dim j As Long
          Dim NR As Long
          Const ErrString = "InputData must have three columns - year, month (as integer 1 to 12), index"
          Dim col_Index
          Dim col_Month
          Dim col_MonthNoFirst
          Dim col_Year
          Dim StartRow As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR InputData
3         If sNCols(InputData) <> 3 Then Throw ErrString
4         NR = sNRows(InputData)
5         If NR < 37 Then Throw "InputData must have at least 37 rows"

6         For i = 1 To sNRows(InputData)
7             For j = 1 To 3
8                 If Not IsNumeric(InputData(i, j)) Then Throw ErrString + ". Non number found at row " + CStr(i) + ", column " + CStr(j)
9             Next
10        Next

11        If InputData(1, 1) <> CLng(InputData(1, 1)) Then Throw ErrString + ". Element 1, 1 is not a whole number"
12        If InputData(1, 2) <> CLng(InputData(1, 2)) Then Throw ErrString + ". Element 1, 2 is not a whole number in the range 1 to 12"
13        If InputData(1, 2) < 1 Or InputData(1, 2) > 12 Then Throw ErrString + ". Element 1, 2 is not a whole number in the range 1 to 12"

14        For i = 2 To NR
15            If InputData(i, 1) <> InputData(i - 1, 1) + IIf(InputData(i - 1, 2) = 12, 1, 0) Then Throw "Error in InputData: Out-of-sequence Year found at row " + CStr(i)
16            If InputData(i, 2) <> InputData(i - 1, 2) + IIf(InputData(i - 1, 2) = 12, -11, 1) Then Throw "Error in InputData: Out-of-sequence Month found at row " + CStr(i)
17        Next i

18        StartRow = 1
19        If UseThisManyMonths <> 0 Then
20            If UseThisManyMonths < 37 Then Throw "UseThisManyMonths must be omited or be at least 37"
21            If NR > UseThisManyMonths Then
22                StartRow = StartRow + NR - UseThisManyMonths
23            End If
24        End If
25        col_Year = sSubArray(InputData, StartRow, 1, , 1)
26        col_Month = sSubArray(InputData, StartRow, 2, , 1)
27        col_Index = sSubArray(InputData, StartRow, 3, , 1)
28        DataUsed = col_Index
29        FirstMonthUsed = InputData(StartRow, 2)
30        col_MonthNoFirst = sDrop(col_Month, 1)

          Dim MonthOverMonth, DataForSortMerge, AMI    'AverageMonthlyIncrements
31        MonthOverMonth = sArrayDivide(sDrop(col_Index, 1), sDrop(col_Index, -1))
32        DataForSortMerge = sArrayRange(col_MonthNoFirst, MonthOverMonth)
33        AMI = sSubArray(sSortMerge(DataForSortMerge, 1, 2, "Average"), 1, 2, , 1)
          'Rebase
34        AMI = sArrayDivide(AMI, sArrayPower(sColumnProduct(AMI), 1 / 12))
          'TakeLogs
35        AMI = sArrayLog(AMI)

          Dim Sum As Double
36        Sum = AMI(1, 1) + AMI(2, 1) + AMI(3, 1) + AMI(4, 1) + AMI(5, 1) + AMI(6, 1) + AMI(7, 1) + AMI(8, 1) + AMI(9, 1) + AMI(10, 1) + AMI(11, 1) + AMI(12, 1)
37        If Abs(Sum) > 0.000000001 Then Throw "Assertion failed. Sum of 12 elements to return should be 0 but instead it is " + CStr(Sum)
38        EstimateSeasonalAdjustments = AMI

39        Exit Function
ErrHandler:
40        EstimateSeasonalAdjustments = "#EstimateSeasonalAdjustments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveSeasonality
' Author    : Philip Swannell
' Date      : 26-Apr-2017
' Purpose   : Remove seasonality from a data series. Series is assumed to be one data point
'             per month for consecutive months. MonthOfFirstData gives the month (1 to 12) of the
'             first (top) element of Data. First element of return is always equal to first element of Data
' -----------------------------------------------------------------------------------------------------------------------
Function RemoveSeasonality(ByVal Data, MonthOfFirstData As Long, ByVal SeasonalAdjustments)

          Dim Adjustments
          Dim i As Long
          Dim j As Long
          Dim NR As Long
          Dim Sum As Double
          Const SAErr = "SeasonalAdjustments must be a 12-row 1-column array of numbers whose sum is 0"

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti Data, SeasonalAdjustments
3         NR = sNRows(Data)
4         If sNCols(Data) <> 1 Then Throw "Data must have one column"
5         If MonthOfFirstData < 1 Or MonthOfFirstData > 12 Then Throw "MonthOfFirstData must be in the range 1 to 12"

6         If sNRows(SeasonalAdjustments) <> 12 Then Throw SAErr + ", but it has " + CStr(sNRows(SeasonalAdjustments)) + " rows"
7         If sNCols(SeasonalAdjustments) <> 1 Then Throw SAErr + ", but it has " + CStr(sNCols(SeasonalAdjustments)) + " columns"

8         For i = 1 To NR
9             If Not IsNumberOrDate(Data(i, 1)) Then Throw "Non number in Data at row " + CStr(i)
10        Next i

11        Sum = 0
12        For i = 1 To 12
13            If Not IsNumberOrDate(SeasonalAdjustments(i, 1)) Then Throw SAErr + ", but element " + CStr(i) + " is not a number"
14            Sum = Sum + SeasonalAdjustments(i, 1)
15        Next i

16        If Abs(Sum) > 0.000000001 Then Throw SAErr + ", but their sum is " + CStr(Sum)

17        Adjustments = sReshape(1, NR, 1)

18        j = MonthOfFirstData Mod 12 + 1
19        For i = 2 To NR
20            If i > 12 Then
21                Adjustments(i, 1) = Adjustments(i - 12, 1)    'Avoid accumulating errors if sum of elements of SeasonalAdjustments is very slightly different from 0
22            Else
23                Adjustments(i, 1) = Adjustments(i - 1, 1) * Exp(SeasonalAdjustments(j, 1))
24                j = j Mod 12 + 1
25            End If
26        Next i

27        RemoveSeasonality = sArrayDivide(Data, Adjustments)

28        Exit Function
ErrHandler:
29        RemoveSeasonality = "#RemoveSeasonality (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HistoricalInflationVol
' Author    : Philip
' Date      : 13-Jun-2017
' Purpose   : calculates the historical vol of the inflation index and of the seasonaly-adjusted inflation index
' -----------------------------------------------------------------------------------------------------------------------
Function HistoricalInflationVol(ByVal InputData As Variant, Optional UseThisManyMonths As Long)

          Dim AdjustedData
          Dim AdjVol
          Dim FirstMonthUsed As Long
          Dim SeasonalAdjustments
          Dim UnadjustedData
          Dim UnAdjVol
1         On Error GoTo ErrHandler
2         SeasonalAdjustments = ThrowIfError(EstimateSeasonalAdjustments(InputData, UseThisManyMonths, UnadjustedData, FirstMonthUsed))
3         AdjustedData = RemoveSeasonality(UnadjustedData, FirstMonthUsed, SeasonalAdjustments)
4         AdjustedData = sArrayDivide(sDrop(AdjustedData, 1), sDrop(AdjustedData, -1))
5         UnadjustedData = sArrayDivide(sDrop(UnadjustedData, 1), sDrop(UnadjustedData, -1))
6         UnAdjVol = Application.WorksheetFunction.StDev_S(UnadjustedData) * Sqr(12)
7         AdjVol = Application.WorksheetFunction.StDev_S(AdjustedData) * Sqr(12)
8         HistoricalInflationVol = sArraySquare("HistoricalVol", UnAdjVol, "HistoricalVolSA", AdjVol)

9         Exit Function
ErrHandler:
10        HistoricalInflationVol = "#HistoricalInflationVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
