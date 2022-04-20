Attribute VB_Name = "modDateSchedule"
Option Explicit
Option Private Module
'Dec 2021 VBA code to generate swap date schedules. Would like to no longer use R and calling Julia currently has high
'overhead (hopefully to be improved).

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DateSchedule
' Author     : Philip Swannell
' Date       : 17-Dec-2021
' Purpose    : Returns the "dates of a swap", i.e. start date and coupon dates as vector (in fact 2-dimensional array
'              with 1 column). This VBA code follows the Julia code of function dateschedule in my XVA package.
' Parameters :
'  StartDate:
'  EndDate  :
'  Frequency: 1 = Annual, 2 = Semi-Annual, 4 = Quarterly
'  BDC      : Allowed values: 'Mod Foll', 'Foll', 'Mod Prec', 'Prec', 'None'. Adjust for weekends only no bank holidays.
'  WhatToReturn: Allowed values: 'Dates', 'StartDates', 'EndDates'
' -----------------------------------------------------------------------------------------------------------------------
Function DateSchedule(StartDate As Long, EndDate As Long, ByVal Frequency As Variant, Optional BDC As String = "Mod Foll", _
          Optional WhatToReturn As String = "Dates")

          Dim End12YM As Double
          Dim i As Long
          Dim NumCoupons As Long
          Dim PeriodLength As Double
          Dim Result() As Long
          Dim Start12YM As Double

1         On Error GoTo ErrHandler
2         If EndDate <= StartDate Then Throw "EndDate must be after StartDate"
3         PeriodLength = 12 / Frequency
4         Start12YM = Year(StartDate) * 12 + Month(StartDate) + Day(StartDate) / 31
5         End12YM = Year(EndDate) * 12 + Month(EndDate) + Day(EndDate) / 31
6         NumCoupons = Ceil((End12YM - Start12YM) / PeriodLength)

7         ReDim Result(1 To NumCoupons + 1, 1 To 1)

8         Result(1, 1) = AdjustDate(CDate(StartDate), BDC)
9         For i = 0 To NumCoupons - 1
10            Result(NumCoupons - i + 1, 1) = CLng(AdjustDate(ToDate3(Year(EndDate), Month(EndDate) - (i * PeriodLength), Day(EndDate)), BDC))
11        Next i

          'It's possible that the short stub at the start is so short (one day) that date adjustment takes the first two dates to the same date. Test for this and fix.
12        If Result(2, 1) = Result(1, 1) Then
              Dim Tmp
13            Tmp = Result
14            ReDim Result(1 To NumCoupons, 1 To 1)
15            Result(1, 1) = Tmp(1, 1)
16            For i = 2 To NumCoupons
17                Result(i, 1) = Tmp(i + 1, 1)
18            Next
19        End If

20        Select Case LCase(WhatToReturn)
              Case "dates"
21                DateSchedule = Result
22            Case "startdates"
23                DateSchedule = sSubArray(Result, 1, , sNRows(Result) - 1)
24            Case "enddates"
25                DateSchedule = sSubArray(Result, 2)
26            Case Else
27                Throw "WhatToReturn must be one of 'Dates', 'StartDates' or 'EndDates'"
28        End Select

29        Exit Function
ErrHandler:
Stop
Resume
30        DateSchedule = "#DateSchedule (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

'Converts `y`, `m`, `d` to a date where if `d > daysinmonth2(y, m)` then the last day of the month is returned.
'If `d < 1` then the function throws an error. `m` can be outside the range `1:12`
Private Function ToDate3(y As Long, m As Long, d As Long)
          Dim dysim As Long
          Dim mmod12 As Long

1         On Error GoTo ErrHandler
2         If d < 1 Then Throw "d must be at least 1"
3         mmod12 = (m - 1) Mod 12 + 1
4         y = y + (m - mmod12) / 12
5         m = mmod12
6         If d <= 28 Then
7             ToDate3 = DateSerial(y, m, d)
8         Else
9             dysim = DaysInMonth(y, m)
10            ToDate3 = DateSerial(y, m, IIf(d < dysim, d, dysim))
11        End If

12        Exit Function
ErrHandler:
13        Throw "#ToDate3 (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function DaysInMonth(y, m)
    DaysInMonth = Day(DateSerial(y, m + 1, 1) - 1)
End Function

Private Function AdjustDate(TheDate As Date, BDC As String) As Long
          Dim AdjustedDate
          Dim d As Long
          Dim dysim As Long
          Dim NudgeFoll As Long
          Dim NudgePrec As Long
          Dim wd As Long

1         On Error GoTo ErrHandler
2         d = Day(TheDate)
3         dysim = DaysInMonth(Year(TheDate), Month(TheDate))
4         wd = ((TheDate - 2) Mod 7) + 1 'Moday = 1, Sunday = 7
5         NudgeFoll = IIf(wd >= 6, 8 - wd, 0)
6         NudgePrec = IIf(wd >= 6, 5 - wd, 0)
7         If BDC = "Mod Foll" Then
8             AdjustedDate = IIf(d + NudgeFoll > dysim, TheDate + NudgePrec, TheDate + NudgeFoll)
9         ElseIf BDC = "Foll" Then
10            AdjustedDate = TheDate + NudgeFoll
11        ElseIf BDC = "Mod Prec" Then
12            AdjustedDate = IIf(d + NudgePrec < 1, TheDate + NudgeFoll, TheDate + NudgePrec)
13        ElseIf BDC = "Prec" Then
14            AdjustedDate = TheDate + NudgePrec
15        ElseIf BDC = "None" Then
16            AdjustedDate = TheDate
17        Else
18            Throw "Business Day Convention '$bdc' not recognised. Allowed values: 'Mod Foll', 'Foll', 'Mod Prec', 'Prec', 'None'"
19        End If
20        AdjustDate = AdjustedDate

21        Exit Function
ErrHandler:
22        Throw "#AdjustDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function Ceil(x As Double)
1         If x = CLng(x) Then
2             Ceil = x
3         Else
4             Ceil = CLng(x + 0.5)
5         End If
End Function


