Attribute VB_Name = "modTest"
Option Explicit

Sub testAddChart()
1         AddXYChart ActiveSheet.Range("XDataWithHeaders"), ActiveSheet.Range("YDataWithHeaders"), "A Title", Range("L10"), 300, 400
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddXYChart
' Author     : Philip Swannell
' Date       : 20-May-2019
' Purpose    : Very basic XY chart. Want to be able to write code to add charts with little effort
' Parameters :
'  XDataWithHeaders:
'  YDataWithHeaders:
'  Title           :
'  ChartTopLeft    :
'  ChartHeight     :
'  ChartWidth      :
' -----------------------------------------------------------------------------------------------------------------------
Sub AddXYChart(XDataWithHeaders As Range, YDataWithHeaders As Range, Title As String, ChartTopLeft As Range, ChartHeight As Long, ChartWidth As Long, Optional xAxisMin = "Auto", Optional xAxisMax = "Auto")
          Dim ch As Chart
          Dim i As Long
          Dim LeftPart As String
          Dim o As Shape
          Dim ws
          Dim XData As Range
          Dim YData As Range
1         On Error GoTo ErrHandler
          'Do input checking here

2         With XDataWithHeaders
3             Set XData = .Offset(1).Resize(.Rows.Count - 1)
4         End With
5         With YDataWithHeaders
6             Set YData = .Offset(1).Resize(.Rows.Count - 1)
7         End With

8         Set ws = ChartTopLeft.Parent
9         Set o = ws.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers)
10        o.Top = ChartTopLeft.Top
11        o.Left = ChartTopLeft.Left
12        o.Height = ChartHeight
13        o.Width = ChartWidth
14        o.Placement = xlMove
15        Set ch = o.Chart
16        For i = 1 To YDataWithHeaders.Columns.Count
17            LeftPart = "='" & ws.Name & "'!"
18            ch.SeriesCollection.NewSeries
19            ch.FullSeriesCollection(i).Name = LeftPart & YDataWithHeaders.Cells(1, i).address
20            ch.FullSeriesCollection(i).xValues = LeftPart & XData.Columns(SafeMin(i, XData.Columns.Count)).address
21            ch.FullSeriesCollection(i).Values = LeftPart & YData.Columns(i).address
22        Next i

23        If Title <> vbNullString Then
24            ch.SetElement (msoElementChartTitleAboveChart)
25            ch.ChartTitle.text = Title
26        Else
27            On Error Resume Next
28            ch.ChartTitle.Delete
29            On Error GoTo ErrHandler
30        End If

31        ch.SetElement (msoElementLegendBottom)

32        If xAxisMin <> "Auto" Then
33            ch.Axes(xlCategory).MinimumScale = xAxisMin
34        End If
35        If xAxisMax <> "Auto" Then
36            ch.Axes(xlCategory).MaximumScale = xAxisMax
37        End If

38        Exit Sub
ErrHandler:
39        Throw "#AddXYChart (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub TestStacker()
          Dim i As Long

          Dim STK As clsStacker

1         Set STK = CreateStacker()

2         For i = 1 To 5000
3             STK.Stack0D "   " & ChrW$(i) & "    "
4         Next

5         g STK.Report

End Sub
