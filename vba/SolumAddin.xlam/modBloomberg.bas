Attribute VB_Name = "modBloomberg"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBBSwaptionVolCode
' Author    : Philip Swannell
' Date      : 08-Jun-2016
' Purpose   : For use with the Bloomberg functions BDH and BDP, returns the string required
'             as the first argument to those functions when we wish to access a point within
'             the ATM swaption vol matrix. We'll need to take care for each currency as
'             to what contributor and what QuoteType to use.
' -----------------------------------------------------------------------------------------------------------------------
Function sBBSwaptionVolCode(Ccy As String, Exercise As String, Tenor As String, Optional QuoteType As String = "Normal", Optional Contributor = "CMPN") As String
1         On Error GoTo ErrHandler
          Dim Part1 As String
          Dim Part2 As String
          Dim Part3 As String
          Dim Part4 As String
          Dim Part5 As String
          Const Part6 = " Curncy"

          'TODO investigate in Bloomberg which other currencies are available...
2         Select Case Ccy
              Case "AUD": Part1 = "AD"
3             Case "CAD": Part1 = "CD"
4             Case "CHF": Part1 = "SF"
5             Case "EUR": Part1 = "EU"
6             Case "GBP": Part1 = "BP"
7             Case "JPY": Part1 = "JY"
8             Case "NZD": Part1 = "ND"
9             Case "USD": Part1 = "US"
10            Case Else: Throw "Unrecognised Ccy. Allowed values are: AUD, CAD, CHF, EUR, GBP, JPY, NZD, USD"
11        End Select

12        Select Case QuoteType
              Case "Normal": Part2 = "SN"
13            Case "Log Normal": Part2 = "SV"
14            Case "OIS Normal": Part2 = "NE"
15            Case "OIS Log Normal": Part2 = "VE"
16            Case Else: Throw "Unrecognised QuoteType. Allowed values are: Normal, Log Normal, OIS Normal, OIS Log Normal"
17        End Select

18        Select Case UCase$(Exercise)
              Case "1Y", "12M"
19                Part3 = "01"
20            Case "1M", "2M", "3M", "4M", "5M", "6M", "7M", "8M", "9M", "10M", "11M"
21                Part3 = "0" & Chr$(64 + CLng(Left$(Exercise, Len(Exercise) - 1)))        '1M = 0A, 2M = 0B etc
22            Case "13M", "14M", "15M", "16M", "17M", "18M"
23                Part3 = "1" & Chr$(64 - 12 + CLng(Left$(Exercise, Len(Exercise) - 1)))        '13M = 1A, 14M = 1B etc
24            Case "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "25Y", "30Y"
25                Part3 = Format$(CLng(Left$(Exercise, Len(Exercise) - 1)), "00")
26            Case Else
27                Throw "Unrecognised Exercise. Allowed values are: 1M, 2M, 3M,...18M, 1Y, 2Y, 3Y, 4Y, 5Y, 6Y, 7Y, 8Y, 9Y, 10Y, 15Y, 20Y, 25Y, 30Y"
28        End Select

29        Select Case UCase$(Tenor)
              Case "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "25Y", "30Y"
30                Part4 = Left$(Tenor, Len(Tenor) - 1)
31            Case Else
32                Throw "Unrecognised Tenor. Allowed values are: 1Y, 2Y, 3Y, 4Y, 5Y, 6Y, 7Y, 8Y, 9Y, 10Y, 15Y, 20Y, 25Y, 30Y"
33        End Select

34        Select Case Contributor
              Case "CMPN", "LAST", "BBIR", "CFIR", "CNTR", "ICPL", "SMKO", "TRPU", "GFIS"
35                Part5 = " " & Contributor
36            Case Else
37                Throw "Unrecognised Contributor. Allowed values are: CMPN, LAST, BBIR, CFIR, CNTR, ICPL, SMKO, TRPU, GFIS"
38        End Select

39        sBBSwaptionVolCode = Part1 & Part2 & Part3 & Part4 & Part5 & Part6

40        Exit Function
ErrHandler:
41        sBBSwaptionVolCode = "#sBBSwaptionVolCode (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
