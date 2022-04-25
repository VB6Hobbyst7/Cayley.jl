Attribute VB_Name = "modWizard"
Option Explicit
Public gApplyRandomAdjustments As Boolean

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedRatesFromBloomberg
' Author    : Philip Swannell
' Date      : 02-Dec-2016
' Purpose   : Wizard style dialog replaces for choosing what to feed from Bloomberg replaces
'             former use of command-bar menu with too many choices.
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedRatesFromBloomberg()
          Dim Live As Boolean
          Const Title = "Feed Rates"
          Const TopText1 = "Live rates or close of business rates?"
          Const TopText2 = "What types of rate do you want to feed?"
          Dim AsOfDate As Long
          Dim ButtonClicked1 As String
          Dim TopText3 As String
          Static FirstChoice As String
          Static SecondChoice As Variant
          Static ThirdChoice As Variant
          Const chLive = "Live Rates"
          Const chCoB = "Close of Business Rates"
          Const chFeedFx = "Fx spot and vol"
          Const chFeedSwapRates = "Swap rates"
          Const chFeedXccyBasis = "Cross currency basis swap rates"
          Const chFeedIRVol = "Interest rate vol"
          Const chFeedCDS = "Credit spreads"
          Const chFeedInflation = "Inflation swaps"
          Const chFeedInflationHistoricSets = "Inflation Historic Sets"
          Dim ButtonClicked As String
          Dim CopyOfErr As String
          Dim DoBasis As Boolean
          Dim DoCredit As Boolean
          Dim DoFx As Boolean
          Dim DoInflation As Boolean
          Dim DoInflationSets As Boolean
          Dim DoSwaps As Boolean
          Dim DoSwaptions As Boolean
          Dim sheetList
          Dim STK As clsStacker
          Dim TheRateCategories
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
GoBack:
2         FirstChoice = ShowOptionButtonDialog(sArrayStack(chLive, chCoB), Title, TopText1, FirstChoice, , , "Apply random adjustments", gApplyRandomAdjustments, , "&Next >", , "&Cancel", ButtonClicked1)
3         If ButtonClicked1 <> "Next >" Then GoTo earlyExit

4         TheRateCategories = sArrayStack(chFeedFx, chFeedSwapRates, chFeedXccyBasis, chFeedIRVol, chFeedCDS, chFeedInflation, chFeedInflationHistoricSets)
5         If FirstChoice = chLive Then
6             Live = True
7         Else
GoBack2:
8             Live = False
9             AsOfDate = GetCOBDate(ButtonClicked)
10            If ButtonClicked = "< Back" Then GoTo GoBack
11            If AsOfDate = 0 Then GoTo earlyExit
12        End If
GoBack3:
13        SecondChoice = ShowMultipleChoiceDialog(TheRateCategories, SecondChoice, Title, TopText2, , , "< &Back", "&Cancel", False, "&Next >", ButtonClicked)
14        If ButtonClicked = "< Back" Then
15            If Live Then
16                GoTo GoBack
17            Else
18                GoTo GoBack2
19            End If
20        End If

21        If sIsErrorString(SecondChoice) Then GoTo earlyExit

22        DoFx = IsNumber(sMatch(chFeedFx, SecondChoice))
23        DoSwaps = IsNumber(sMatch(chFeedSwapRates, SecondChoice))
24        DoBasis = IsNumber(sMatch(chFeedXccyBasis, SecondChoice))
25        DoSwaptions = IsNumber(sMatch(chFeedIRVol, SecondChoice))
26        DoCredit = IsNumber(sMatch(chFeedCDS, SecondChoice))
27        DoInflation = IsNumber(sMatch(chFeedInflation, SecondChoice))
28        DoInflationSets = IsNumber(sMatch(chFeedInflationHistoricSets, SecondChoice))

29        TopText3 = "Feed " + IIf(Live, "live rates for:", "close of business rates for " + Format(AsOfDate, "d-mmm-yyyy")) + vbLf + _
              sConcatenateStrings(SecondChoice, vbLf) + vbLf + vbLf + "Choose currencies"

30        Set STK = CreateStacker()

31        For Each ws In ThisWorkbook.Worksheets
32            If IsCurrencySheet(ws) Then
33                STK.StackData ws.Name
34            End If
35        Next ws

36        If DoSwaps Or DoBasis Or DoSwaptions Then
37            sheetList = sSortedArray(STK.report)
38            ThirdChoice = ShowMultipleChoiceDialog(sheetList, ThirdChoice, Title, TopText3, , , "< &Back", "&Cancel", False, "&OK", ButtonClicked)
39            If ButtonClicked = "< Back" Then GoTo GoBack3
40            If sIsErrorString(ThirdChoice) Then GoTo earlyExit
41        Else
42            ThirdChoice = Empty
43        End If

44        FeedAllRatesAllSheets Live, AsOfDate, ThirdChoice, DoFx, DoSwaps, DoBasis, DoSwaptions, DoCredit, DoInflation, DoInflationSets

earlyExit:
45        Application.Cursor = xlDefault

46        Exit Sub
ErrHandler:
47        CopyOfErr = "#FeedRatesFromBloomberg (line " & CStr(Erl) + "): " & Err.Description & "!"
48        Application.Cursor = xlDefault
49        SomethingWentWrong CopyOfErr
End Sub



