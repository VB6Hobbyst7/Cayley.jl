Attribute VB_Name = "modCharts"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UpdateChartOnCreditUsageSheet
' Author    : Philip Swannell
' Date      : 21-Feb-2017
' Purpose   : "Updates" the chart by deleting it and creating a new one. Earlier version of this method
'             updated the chart but found that this lead to memory leaks. :-(
' -----------------------------------------------------------------------------------------------------------------------
Sub UpdateChartOnCreditUsageSheet(ChartTitle As String, YAxisTitle, BankIsGood As Boolean)
          Dim cht As Chart
          Dim chtOb As ChartObject
          Dim i As Long
          Dim Sh As Shape
1         On Error GoTo ErrHandler

          'Try to avoid recreating the chart, which gives quite a bit of screen flicker, just refresh it
2         If shCreditUsage.ChartObjects.Count = 1 Then
3             Set cht = shCreditUsage.ChartObjects(1).Chart
4             If cht.FullSeriesCollection.Count = IIf(BankIsGood, 2, 1) Then
5                 If cht.Axes(xlCategory).MaximumScale = GetHedgeHorizon() + 1 Then
6                     If UBound(cht.FullSeriesCollection(1).xValues) = _
                          RangeFromSheet(shCreditUsage, "TheData").Rows.Count Then
                      
7                         cht.Parent.Visible = True
8                         If cht.ChartTitle.Caption <> ChartTitle Then
9                             cht.ChartTitle.Caption = ChartTitle
10                        End If
11                        If cht.Axes(xlValue).DisplayUnitLabel.Caption <> YAxisTitle Then
12                            cht.Axes(xlValue).DisplayUnitLabel.Caption = YAxisTitle
13                        End If

                          'Mmmm cht.Refresh does not seem to do what it says on the tin! _
                           Have to recalc the sheet but control EnableEvents to prevent infinite loop
14                        cht.Refresh
                        
                          Dim OldEE As Boolean
15                        OldEE = Application.EnableEvents
16                        If OldEE Then Application.EnableEvents = False
17                        shCreditUsage.Calculate
18                        If OldEE Then Application.EnableEvents = True
19                        Exit Sub
20                    End If
21                End If
22            End If
23        End If

24        For Each chtOb In shCreditUsage.ChartObjects
25            chtOb.Delete
26        Next

27        If Val(Application.Version) > 14 Then
              'Office 2013 and later
28            Set Sh = shCreditUsage.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers)
29        Else
              'Office 2010
30            Set Sh = shCreditUsage.Shapes.AddChart(xlXYScatterLinesNoMarkers)
31        End If

32        Set cht = Sh.Chart
33        cht.PlotVisibleOnly = False
34        cht.Parent.Visible = True

35        For i = cht.SeriesCollection.Count To 1 Step -1
36            cht.SeriesCollection(i).Delete
37        Next i
38        With cht.SeriesCollection.NewSeries
39            .xValues = "=" & shCreditUsage.Name & "!" & RangeFromSheet(shCreditUsage, "TheData").Columns(2).Address
40            .Values = "=" & shCreditUsage.Name & "!" & RangeFromSheet(shCreditUsage, "TheData").Columns(3).Address
41            .Name = "=" & shCreditUsage.Name & "!" & RangeFromSheet(shCreditUsage, "TheData").Cells(0, 3).Address
42        End With
43        If BankIsGood Then
44            With cht.SeriesCollection.NewSeries
45                .xValues = "=" & shCreditUsage.Name & "!" & _
                      RangeFromSheet(shCreditUsage, "CreditLimitsForPlotting").Columns(1).Address
46                .Values = "=" & shCreditUsage.Name & "!" & _
                      RangeFromSheet(shCreditUsage, "CreditLimitsForPlotting").Columns(2).Address
47                .Name = "Line"
48            End With
49        End If

50        cht.Axes(xlCategory).TickLabels.NumberFormat = "0"
51        cht.Axes(xlCategory).MaximumScale = GetHedgeHorizon() + 1

52        cht.Axes(xlValue).DisplayUnit = xlMillions
53        cht.SetElement (msoElementChartTitleAboveChart)
54        cht.ChartTitle.Caption = ChartTitle

55        With cht.ChartTitle.Format.TextFrame2.TextRange.Font
56            .Fill.ForeColor.RGB = RGB(87, 87, 87)
57            .Fill.Transparency = 0
58            .Size = 14
59            .Bold = msoFalse
60        End With

61        cht.Axes(xlValue).DisplayUnitLabel.Caption = YAxisTitle
62        cht.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
63        cht.SetElement (msoElementLegendBottom)
64        cht.SetElement (msoElementPrimaryCategoryGridLinesMajor)

65        cht.Axes(xlCategory).AxisTitle.Caption = "Time (years)"

66        PositionChartOnCreditUsageSheet

67        Exit Sub
ErrHandler:
68        Throw "#UpdateChartOnCreditUsageSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub PositionChartOnCreditUsageSheet()

          Dim BR As Range
          Dim TL As Range
          Dim Target As Range

1         On Error GoTo ErrHandler

2         With RangeFromSheet(shCreditUsage, "ExtraTradeAmounts")
3             Set TL = .Cells(.Rows.Count + 2, 0)
4         End With
              
5         Set Target = Range(TL, TL.offset(22, 8))

6         With shCreditUsage.ChartObjects(1)
7             .Top = Target.Top
8             .Left = Target.Left
9             .Width = Target.Width
10            .Height = Target.Height
11        End With

12        With RangeFromSheet(shCreditUsage, "FilterBy1").offset(0, 2).Resize(, 2).EntireColumn
13            .Hidden = False
              'so that the column doesn't "pop-back to life" when entering a formula and _
               selecting cells to be part of that formula.
14            .ColumnWidth = 0.05
15            .Hidden = True
16        End With

17        Exit Sub
ErrHandler:
18        Throw "#PositionChartOnCreditUsageSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExpandBubbleChartButton
' Author    : Philip Swannell
' Date      : 01-Nov-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub ExpandBubbleChartButton()
          Dim b As Button
          Dim Expand As Boolean
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shBubbleChart)
3         Set SUH = CreateScreenUpdateHandler

4         Set b = shBubbleChart.Buttons(Application.Caller)

5         If b.Caption = "z" Then Expand = True
6         ExpandChart2 Expand, shBubbleChart, b
7         Set SPH = Nothing

8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#ExpandBubbleChartButton (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExpandBarChartButton
' Author    : Philip Swannell
' Date      : 01-Nov-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub ExpandBarChartButton()
          Dim b As Button
          Dim Expand As Boolean
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shBarChart)
3         Set SUH = CreateScreenUpdateHandler

4         Set b = shBarChart.Buttons(Application.Caller)

5         If b.Caption = "z" Then Expand = True
6         ExpandChart2 Expand, shBarChart, b
7         Set SPH = Nothing

8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#ExpandBarChartButton (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExpandChart2
' Author    : Philip Swannell
' Date      : 17-Oct-2016
' Purpose   : Version of ExpandChart for use on the BarChart and BubbleChart sheets
' -----------------------------------------------------------------------------------------------------------------------
Sub ExpandChart2(Expand As Boolean, ws As Worksheet, b As Button)
          Dim co As ChartObject
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set co = ws.ChartObjects(1)
3         Set SPH = CreateSheetProtectionHandler(ws)
          Dim H As Double
          Dim l As Double
          Dim T As Double
          Dim W As Double

4         If Expand Then
5             T = 47.25: l = 24: W = 652: H = 307
6         Else
7             T = 47.25: l = 24: W = 652 * 1.8: H = 307 * 1.8
8         End If

9         co.Top = T: co.Left = l: co.Width = W: co.Height = H

10        If Expand Then
11            b.Caption = "y"
12        Else
13            b.Caption = "z"
14        End If

15        With b
16            .Placement = xlMove
17            .Width = 15
18            .Height = 15
19            .Top = T
20            .Left = l
21            .Font.ColorIndex = 48
22        End With
23        Exit Sub
ErrHandler:
24        Throw "#ExpandChart2 (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SummariseFilters
' Author    : Philip Swannell
' Date      : 20-May-2015
' Purpose   : Translate the filters into natural language
' -----------------------------------------------------------------------------------------------------------------------
Function SummariseFilters(FilterBy1 As String, Filter1Value, FilterBy2 As String, Filter2Value)
          Dim NullFirstFilter As Boolean
          Dim NullSecondFilter As Boolean
          Dim Result As String

1         On Error GoTo ErrHandler
2         NullFirstFilter = LCase(FilterBy1) = "none" Or LCase(Filter1Value) = "all"
3         NullSecondFilter = LCase(FilterBy2) = "none" Or LCase(Filter2Value) = "all"

4         If NullFirstFilter And NullSecondFilter Then
5             Result = "All trades"
6         ElseIf Not NullFirstFilter Then
7             If FilterBy1 = "Counterparty Parent" Then
8                 Result = "Trades with " & _
                      FirstElementOf(LookupCounterpartyInfo(Filter1Value, "CPTY LONG NAME", Filter1Value, Filter1Value))
9             Else
10                Result = "Trades where '" & FilterBy1 & "' matches '" & Abbreviate(CStr(Filter1Value), 30) & "'"
11            End If
12        End If
13        If Not NullSecondFilter Then
14            If Not NullFirstFilter Then
15                Result = Result & " and "
16            Else
17                Result = "Trades with "
18            End If
19            Result = Result & "'" & CStr(FilterBy2) & "' matches '" & Abbreviate(CStr(Filter2Value), 30) & "'"
20        End If
21        SummariseFilters = Result

22        Exit Function
ErrHandler:
23        SummariseFilters = "#SummariseFilters (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Abbreviate
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : sub-routine of SummariseFilters needed since the regular expressions can be
'             way too long to appear in the graph title.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Abbreviate(AString As String, MaxLen As Long) As String
1         On Error GoTo ErrHandler
2         If Len(AString) > MaxLen Then
3             Abbreviate = Left(AString, MaxLen - 8) & "..." & Right(AString, 5)
4         Else
5             Abbreviate = AString
6         End If
7         Exit Function
ErrHandler:
8         Throw "#Abbreviate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PFEChartTitle
' Author    : Philip Swannell
' Date      : 20-May-2015
' Purpose   : Encapsulate generating a helpful title for the chart on sheet CreditUsage
' -----------------------------------------------------------------------------------------------------------------------
Function PFEChartTitle(FilterBy1 As String, Filter1Value, FilterBy2 As String, Filter2Value, IncludeExtraTrades, _
          ExtraTradeAmounts, PortfolioAgeing As Double, FxShock, FxVolShock, TradesScaleFactor As Double, _
          LinesScaleFactor As Double, ByVal NumTrades As Long, BankIsGood As Boolean, IncludeFxTrades As Boolean, _
          IncludeRatesTrades As Boolean, ExtraMessage As String)

          Dim NumExtraTrades As Long
          Dim Result

1         On Error GoTo ErrHandler
2         Result = SummariseFilters(FilterBy1, Filter1Value, FilterBy2, Filter2Value)
3         Result = Result & vbLf & Format(NumTrades, "###,##0") & " trade" & IIf(NumTrades <> 1, "s", "")
4         If IncludeExtraTrades Then
5             NumExtraTrades = sArrayCount(sArrayGreaterThan(sArrayAbs(ExtraTradeAmounts), 0))
6         End If
7         If NumExtraTrades > 0 Then
8             Result = Result & " plus " & CStr(NumExtraTrades) & " extra"
9         End If

10        If PortfolioAgeing > 0 Then
11            Result = Result & ", Trades aged by " & PortfolioAgingToString(CDbl(PortfolioAgeing))
12        ElseIf PortfolioAgeing < 0 Then
13            Result = Result & ", Trades shifted forward by " & PortfolioAgingToString(-CDbl(PortfolioAgeing))
14        End If

15        If FxShock < 1 Then
16            Result = Result & ", EUR down " & Format(1 - FxShock, "0%")
17        ElseIf FxShock > 1 Then
18            Result = Result & ", EUR up " & Format(FxShock - 1, "0%")
19        End If

20        If FxVolShock < 1 Then
21            Result = Result & ", Fx Vol down " & Format(1 - FxVolShock, "0%")
22        ElseIf FxVolShock > 1 Then
23            Result = Result & ", Fx Vol up " & Format(FxVolShock - 1, "0%")
24        End If

25        If TradesScaleFactor <> 1 Then
26            If NumTrades <> 0 Then
27                Result = Result & ", Trades scaled " & CStr(TradesScaleFactor)
28            End If
29        End If

30        If BankIsGood Then
31            If LinesScaleFactor > 1 Then
32                Result = Result & ", Lines up " & Format(LinesScaleFactor - 1, "0%")
33            ElseIf LinesScaleFactor < 1 Then
34                Result = Result & ", Lines down " & Format(1 - LinesScaleFactor, "0%")
35            End If
36        End If

37        Result = Trim(Replace(Result, "  ", " "))
38        If Right(Result, 1) <> "." Then
39            Result = Result & "."
40        End If

41        If IncludeRatesTrades And IncludeFxTrades Then
42            Result = "Rates and Fx " & Result
43        ElseIf IncludeRatesTrades Then
44            Result = "Rates " & Result
45        ElseIf IncludeFxTrades Then
46            Result = "Fx " & Result
47        End If

48        If ExtraMessage <> "" Then Result = Result & " " & ExtraMessage

49        PFEChartTitle = Result

50        Exit Function
ErrHandler:
51        PFEChartTitle = "#PFEChartTitle (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function RoundDown(x As Double) As Long
1         If x = CLng(x) Then
2             RoundDown = CLng(x)
3         Else
4             RoundDown = CLng(x - 0.5)
5         End If
End Function

Function PortfolioAgingToString(PA As Double) As String

          Dim ND As Long
          Dim NDstr As String
          Dim NM As Long
          Dim NMstr As String
          Dim NY As Long
          Dim NYstr As String

1         On Error GoTo ErrHandler
2         If PA > 0 Then
3             NY = RoundDown(PA)
4             NM = RoundDown((PA - NY) * 12)
5             ND = RoundDown((PA - NY - NM / 12) * 360)
6             NYstr = IIf(NY = 0, "", CStr(NY) & " year") & IIf(NY > 1, "s", "") & IIf(NY = 0, "", " ")
7             NMstr = IIf(NM = 0, "", CStr(NM) & " month") & IIf(NM > 1, "s", "") & IIf(NM = 0, "", " ")
8             NDstr = IIf(ND = 0, "", CStr(ND) & " day") & IIf(ND > 1, "s", "") & IIf(ND = 0, "", " ")
9             PortfolioAgingToString = NYstr & NMstr & NDstr
10        ElseIf PA < 0 Then
11            PortfolioAgingToString = PortfolioAgingToString(-PA)
12        End If

13        Exit Function
ErrHandler:
14        Throw "#PortfolioAgingToString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnselectChart
' Author    : Philip Swannell
' Date      : 14-Oct-2016
' Purpose   : User has some level of control over Excel while the macro is running and it's
'             all too easy to select the chart, if they have unselect it.
' -----------------------------------------------------------------------------------------------------------------------
Sub UnselectChart()
1         On Error GoTo ErrHandler
2         If TypeName(Selection) <> "Range" Then
              'user can accidentally select the graph while macro is running
3             ActiveWindow.RangeSelection.Select
4         End If
5         Exit Sub
ErrHandler:
6         Throw "#UnselectChart (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CopyChart
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Copies chart from PFE sheet to a new location.
'             Take care - this does a break links so the target needs to be in a new workbook!
' -----------------------------------------------------------------------------------------------------------------------
Sub CopyChart(co As ChartObject, Target As Range)
1         On Error GoTo ErrHandler

2         co.Chart.ChartArea.Copy
3         Application.GoTo Target

4         Target.Parent.Paste
5         Target.Parent.Parent.BreakLink Name:=co.Parent.Parent.FullName, _
              Type:=xlExcelLinks
6         Target.Select
7         Exit Sub
ErrHandler:
8         Throw "#CopyChart (line " & CStr(Erl) & "): " & Err.Description & "!"
9     End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PrintCharts
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Automate printing of many charts
' -----------------------------------------------------------------------------------------------------------------------
Sub PrintCharts(Optional ToPaper As Boolean = False, Optional BanksToProcess As Variant, Optional TargetFolder As String, _
          Optional ExportJPG As Boolean, Optional AnchorDate As Date, Optional SilentMode As Boolean = False)
          
          Dim Title As String

          Dim AllCounterparties
          Dim BookFullName As String
          Dim BookName As String
          Dim c
          Dim CurrenciesToInclude As String
          Dim Filter2Value As Variant
          Dim FilterBy2 As String
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim IncludeAssetClasses As String
          Dim IncludeFutureTrades As Boolean
          Dim JPGName As String
          Dim LinesBook As Workbook
          Dim LinesScaleFactor As Double
          Dim NamesNotPrinted
          Dim NamesPrinted As String
          Dim NumMCPaths As Long
          Dim NumNotPrinted As Long
          Dim NumObservations As Long
          Dim NumPrinted As Long
          Dim PortfolioAgeing As Double
          Dim Prompt As String
          Dim SPH As Object
          Dim Suffix As String
          Dim SUH As Object
          Dim TradesScaleFactor As Double

1         On Error GoTo ErrHandler

2         If ToPaper Then
3             Title = "Print Charts"
4         Else
5             Title = "Paste Charts"
6         End If

7         Set SPH = CreateSheetProtectionHandler(shCreditUsage)

8         If IsEmpty(BanksToProcess) Or IsMissing(BanksToProcess) Then
9             Set LinesBook = OpenLinesWorkbook(True, False)
10            AllCounterparties = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
11            AllCounterparties = sSortedArray(AllCounterparties)
12            AllCounterparties = AnnotateBankNames(AllCounterparties, True, LinesBook)
13            BanksToProcess = ShowMultipleChoiceDialog(AllCounterparties, , Title, _
                  "Select Parent Counterparties for which to " & IIf(ToPaper, "print", "paste") & " charts.", False)
14            If VarType(BanksToProcess) < vbArray Then GoTo EarlyExit
15            BanksToProcess = AnnotateBankNames(BanksToProcess, False, LinesBook)
16        End If

17        FilterBy2 = RangeFromSheet(shCreditUsage, "FilterBy2", False, True, False, False, False)
18        Filter2Value = RangeFromSheet(shCreditUsage, "Filter2Value", True, True, True, False, False)
19        IncludeFutureTrades = RangeFromSheet(shCreditUsage, "IncludeFutureTrades", False, False, True, False, False)
20        IncludeAssetClasses = RangeFromSheet(shCreditUsage, "IncludeAssetClasses", False, True, False, False, False)
21        PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False)
22        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False)
23        CurrenciesToInclude = RangeFromSheet(shConfig, "CurrenciesToInclude", False, True, False, False, False)
24        NumMCPaths = RangeFromSheet(shCreditUsage, "NumMCPaths", True, False, False, False, False)
25        NumObservations = RangeFromSheet(shCreditUsage, "NumObservations", True, False, False, False, False)
26        FxShock = RangeFromSheet(shCreditUsage, "FxShock", True, False, False, False, False)
27        FxVolShock = RangeFromSheet(shCreditUsage, "FxVolShock", True, False, False, False, False)
28        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False)

29        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value
30        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False).Value
31        shCreditUsage.Activate

          Dim PromptArray
          Dim PromptArrayShort

32        PortfolioAgeing = RangeFromSheet(shCreditUsage, "PortfolioAgeing", True, False, False, False, False).Value

33        OpenOtherBooks
34        JuliaLaunchForCayley
35        BuildModelsInJulia False, FxShock, FxVolShock

36        If Not SilentMode Then
37            If PortfolioAgeing <> 0 Then
38                PromptArray = sArrayStack("PortfolioAgeing", _
                      RangeFromSheet(shCreditUsage, "PortfolioAgeing").Value, _
                      "Portfolio aged to", _
                      Format(gModel_CM("AnchorDate") + _
                      RangeFromSheet(shCreditUsage, "PortfolioAgeing") * 365, "dd-mmm-yyyy"))

39            Else
40                PromptArray = CreateMissing()
41            End If
42            PromptArray = sArrayStack(PromptArray, _
                  "NumMCPaths", NumMCPaths, _
                  "NumObservations", NumObservations, _
                  "IncludeFutureTrades", IncludeFutureTrades, _
                  "FilterBy2", FilterBy2, _
                  "Filter2Value", Filter2Value, _
                  "CurrenciesToInclude", CurrenciesToInclude, _
                  "FxShock", FxShock, _
                  "FxVolShock", FxVolShock, _
                  "TradesScaleFactor", TradesScaleFactor, _
                  "LinesScaleFactor", LinesScaleFactor)

43            PromptArray = sReshape(PromptArray, sNRows(PromptArray) / 2, 2)
44            PromptArrayShort = CleanUpPromptArray(PromptArray, True)

45            Prompt = IIf(ToPaper, "Print ", "Paste ") & CStr(sNRows(BanksToProcess)) & " chart" & _
                  IIf(sNRows(BanksToProcess) > 1, "s", "") & IIf(ToPaper, " to the printer", " to a new workbook") & "?" & _
                  vbLf & vbLf & _
                  "Charts will be " & IIf(ToPaper, "printed", "pasted") & _
                  " only for banks for which we have good data in the lines workbook and other inputs are as follows:" _
                  & vbLf & sConcatenateStrings(sJustifyArrayOfStrings(PromptArrayShort, "Calibri", 11, vbTab), vbLf)

46            If MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion + vbDefaultButton2, Title) <> vbYes Then GoTo EarlyExit
47        End If

48        g_StartRunCreditUsageSheet = sElapsedTime()

49        Set SUH = CreateScreenUpdateHandler()

50        If ToPaper Then 'In this case don't support exporting as JPG
51            For Each c In BanksToProcess
52                If PrintAChart(CStr(c)) Then
53                    NumPrinted = NumPrinted + 1
54                    NamesPrinted = IIf(NamesPrinted = "", CStr(c), NamesPrinted & ", " & CStr(c))
55                Else
56                    NumNotPrinted = NumNotPrinted + 1
57                    NamesNotPrinted = IIf(NamesNotPrinted = "", CStr(c), NamesNotPrinted & ", " & CStr(c))
58                End If
59            Next
60        Else
              Dim RequiredOffset As Double
              Dim Target As Range
              Dim TargetBook As Workbook
              Dim TargetSheet As Worksheet
61            Set TargetBook = Application.Workbooks.Add
62            Set TargetSheet = TargetBook.Worksheets(1)

63            Suffix = "_" & Format(AnchorDate, "yyyy-mm-dd") & ".jpg"

              'Add headers here
64            With TargetSheet.Cells(1, 1)
65                .Value = "PFE Charts"
66                .Font.Size = 22
67            End With
68            TargetSheet.Cells(2, 1).Value = "Time generated"
69            With TargetSheet.Cells(2, 2)
70                .Value = Now()
71                .NumberFormat = "dd-mmm-yyyy hh:mm"
72                .HorizontalAlignment = xlHAlignLeft
73            End With
74            With TargetSheet.Cells(3, 1).Resize(sNRows(PromptArray), 2)
75                .Value = sArrayExcelString(PromptArray)
76                .HorizontalAlignment = xlHAlignLeft
77            End With
78            TargetSheet.UsedRange.Columns.AutoFit
79            TargetBook.Windows(1).DisplayGridlines = False
80            TargetBook.Windows(1).DisplayHeadings = False

81            Set Target = TargetSheet.Cells(TargetSheet.UsedRange.Rows.Count + 2, 1)

              Dim i As Long
              Dim NumBanks As Long
82            NumBanks = sNRows(BanksToProcess)
83            i = 0
84            For Each c In BanksToProcess
85                i = i + 1
86                StatusBarWrap "Generating chart " & CStr(i) & " of " & CStr(NumBanks) & " " & CStr(c)
87                PrepareForCalculation c, False, False, True
88                RunCreditUsageSheet "Standard", True, False, True
89                CopyChart shCreditUsage.ChartObjects(1), Target
90                If ExportJPG Then
91                    JPGName = sJoinPath(TargetFolder, c & Suffix)
92                    shCreditUsage.ChartObjects(1).Chart.Export JPGName
93                End If

94                If RequiredOffset = 0 Then
                      'Need to calculate this quantity inside the loop. No guarantee that there is a chart on the sheet before the loop starts
95                    RequiredOffset = CLng(shCreditUsage.ChartObjects(1).Height / 14.25) + 1
96                End If
97                Set Target = Target.offset(RequiredOffset)
98            Next c
99            StatusBarWrap False
100       End If
101       FormatCreditUsageSheet True

102       If Not SilentMode Then
103           If ToPaper Then
104               Application.GoTo shCreditUsage.Cells(1, 1)
105           Else
106               Application.GoTo TargetSheet.Cells(1, 1)
107           End If
108       End If

109       If TargetFolder <> "" Then
110           BookName = "AllPFECharts_" & Format(AnchorDate, "yyyy-mm-dd") & ".xlsx"
111           BookFullName = sJoinPath(TargetFolder, BookName)
112           If IsInCollection(Application.Workbooks, BookName) Then
113               Application.Workbooks(BookName).Close False
114           End If
115           If sFileExists(BookFullName) Then
116               ThrowIfError sFileDelete(BookFullName)
117           End If
118           TargetBook.SaveAs BookFullName, xlOpenXMLWorkbook
119           TargetBook.Close False
120           ThisWorkbook.Activate
121       End If

122       If Not SilentMode Then
123           If NumNotPrinted > 0 Then
124               Prompt = CStr(NumPrinted) & " of " & CStr(sNRows(BanksToProcess)) & " charts were printed." & vbLf & vbLf & _
                      "The following charts were not printed because headroom could not be calculated, " & _
                      "probably because of bad data in the lines workbook" & vbLf & NamesNotPrinted
125               If gDebugMode Then Debug.Print Prompt
126               MsgBoxPlus Prompt, vbOKOnly + vbExclamation, Title
127           Else
128               Prompt = "All " + CStr(sNRows(BanksToProcess)) & " chart(s) were printed."
129               If ToPaper Then
130                   MsgBoxPlus Prompt, vbOKOnly + vbInformation, Title
131               End If
132           End If
133       End If

EarlyExit:
134       Set TargetSheet = Nothing
135       Set TargetBook = Nothing
136       Set LinesBook = Nothing

137       Exit Sub
ErrHandler:
138       SomethingWentWrong "#PrintCharts (line " & CStr(Erl) & "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PrintAChart
' Author    : Philip Swannell
' Date      : 06-Oct-2016
' Purpose   : Given a Parent Counterparty, refreshes the chart on the PFE sheet (with other
'             inputs to the sheet in their current state) and prints out a chart.
' -----------------------------------------------------------------------------------------------------------------------
Function PrintAChart(Counterparty As String)

1         On Error GoTo ErrHandler
2         PrepareForCalculation Counterparty, False, False, True
3         RunCreditUsageSheet "Standard", True, False, True
4         shCreditUsage.ChartObjects(1).Activate

5         If IsNumeric(RangeFromSheet(shCreditUsage, "MaxPFEByYear").Cells(1, 1)) Then
6             ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                  IgnorePrintAreas:=False
7             PrintAChart = True
8         Else
9             PrintAChart = False
10        End If

11        Exit Function
ErrHandler:
12        Throw "#PrintAChart (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function



