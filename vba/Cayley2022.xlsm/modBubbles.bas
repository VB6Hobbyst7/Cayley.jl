Attribute VB_Name = "modBubbles"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NineBubbleCharts
' Author    : Philip Swannell
' Date      : 08-Nov-2016
' Purpose   : This method runs the table nine times to create nine 'Bubble Charts'. The runs
'             use FxShocks of 0.9, 1 and 1.1 and FxVolShocks of 0.9, 1 and 1.1.
'             Charts are pasted to a new workbook and the method may take some hours to run.
'             Attached to Menu on Table sheet.
' -----------------------------------------------------------------------------------------------------------------------
Sub NineBubbleCharts()
          Dim BanksToRun
          Dim FileName As String
          Dim FxShock As Double
          Dim FxVolShock As Double
          Dim i As Long
          Dim Prompt As String
          Dim wb As Workbook
          Const Title = "Nine Bubble Charts"

1         On Error GoTo ErrHandler

2         Prompt = "This method runs the table nine times to create nine 'Bubble Charts'. The runs use " & _
              "FxShocks of 0.9, 1 and 1.1 and FxVolShocks of 0.9, 1 and 1.1." & vbLf & vbLf & _
              "Charts are pasted to a new workbook which is saved to the ScenarioResultsDirectory specified " & _
              "on the Config sheet." & vbLf & vbLf & _
              "The method may take some hours to run." & vbLf & vbLf & "Proceed?"

3         If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, "Yes, Run", "No, don't run") <> vbOK Then Exit Sub

4         FileName = RangeFromSheet(shConfig, "ScenarioResultsDirectory")
5         If Right(FileName, 1) <> "\" Then FileName = FileName & "\"
          Dim Res

6         Res = sCreateFolder(FileName)
7         If sIsErrorString(Res) Then
8             Throw "Error when attempting to create a directory to save the workbook containing charts: " & _
                  Res & vbLf & _
                  "Please check that 'ScenarioResultsDirectory' on the Config sheet is set to " & _
                  "a location to which you have write access."
9         End If
10        If Not sFolderIsWritable(FileName) Then
11            Throw "You do not have write access to the scenario results folder (also used for the " & _
                  "workbook containing nine bubble charts)." & vbLf & _
                  "Please check that 'ScenarioResultsDirectory' on the Config sheet is set " & _
                  "to a location to which you have write access."
12        End If

13        FileName = FileName & "NineBubbleCharts " & Format(Now, "yyyy-mm-dd hh-mm") & ".xlsx"

14        Application.ScreenUpdating = False
15        Set wb = Application.Workbooks.Add

16        wb.SaveAs FileName

17        For FxVolShock = 0.9 To 1.1 Step 0.1
18            For FxShock = 0.9 To 1.1 Step 0.1
19                i = i + 1
20                MessageLogWrite "NineBubbleCharts starting run " & CStr(i)
21                RangeFromSheet(shCreditUsage, "FxShock").Value = FxShock
22                RangeFromSheet(shCreditUsage, "FxVolShock").Value = FxVolShock
23                With RangeFromSheet(shTable, "TheTable")
24                    BanksToRun = .offset(1).Resize(.Rows.Count - 1, 1).Value
25                End With
26                RunTable True, True, True, BanksToRun
27                SpawnBubbleSheet wb, "Spot x " & CStr(FxShock) & " Vol x " & CStr(FxVolShock)
28                wb.Save
29            Next
30        Next
31        Application.DisplayAlerts = False
32        wb.Worksheets("Sheet1").Delete
33        RescaleAllChartsInBook wb
34        wb.Save

35        Exit Sub
ErrHandler:
36        SomethingWentWrong "#NineBubbleCharts (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub SpawnBubbleSheet(TargetBook As Workbook, TargetSheetName As String)
          Dim SourceRange As Range
          Dim TargetRange As Range
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         shBubbleChart.Calculate
3         Set ws = TargetBook.Worksheets.Add
4         CopyChart shBubbleChart.ChartObjects(1), ws.Cells(1, 1)
5         ws.Name = TargetSheetName
6         ws.Parent.Windows(1).DisplayGridlines = False
7         ws.Parent.Windows(1).DisplayHeadings = False

8         Set SourceRange = RangeFromSheet(shTable, "TheTable")
9         Set TargetRange = ws.Cells(1, 33).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)

10        SourceRange.Copy TargetRange
11        ws.Names.Add "TheTable", TargetRange
12        TargetRange.Columns.AutoFit
13        ws.Protect , True, True

14        Exit Sub
ErrHandler:
15        Throw "#SpawnBubbleSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RescaleAllChartsInBook
' Author    : Philip Swannell
' Date      : 24-Feb-2017
' Purpose   : Puts ALL charts in a workbook onto the same maximum and minimum values for axes scales
' -----------------------------------------------------------------------------------------------------------------------
Sub RescaleAllChartsInBook(wb As Workbook)
          Dim chOb As ChartObject
          Dim isFirst As Boolean
          Dim SPH As clsSheetProtectionHandler
          Dim ThisxMax As Double
          Dim ThisxMin As Double
          Dim ThisyMax As Double
          Dim ThisyMin As Double
          Dim ws As Worksheet
          Dim xMax As Double
          Dim xMin As Double
          Dim yMax As Double
          Dim yMin As Double

1         On Error GoTo ErrHandler
2         isFirst = True

3         For Each ws In wb.Worksheets
4             For Each chOb In ws.ChartObjects
5                 ThisyMax = chOb.Chart.Axes(xlValue).MaximumScale
6                 ThisyMin = chOb.Chart.Axes(xlValue).MinimumScale
7                 ThisxMax = chOb.Chart.Axes(xlCategory).MaximumScale
8                 ThisxMin = chOb.Chart.Axes(xlCategory).MinimumScale
9                 If isFirst Then
10                    yMax = ThisyMax
11                    yMin = ThisyMin
12                    xMax = ThisxMax
13                    xMin = ThisxMin
14                    isFirst = False
15                Else
16                    If yMax < ThisyMax Then yMax = ThisyMax
17                    If yMin > ThisyMin Then yMin = ThisyMin
18                    If xMax < ThisxMax Then xMax = ThisxMax
19                    If xMin > ThisxMin Then xMin = ThisxMin
20                End If
21            Next chOb
22        Next ws

23        For Each ws In wb.Worksheets
24            Set SPH = CreateSheetProtectionHandler(ws)
25            For Each chOb In ws.ChartObjects
26                If chOb.Chart.Axes(xlValue).MaximumScale <> yMax Then
27                    chOb.Chart.Axes(xlValue).MaximumScale = yMax
28                End If
29                If chOb.Chart.Axes(xlValue).MinimumScale <> yMin Then
30                    chOb.Chart.Axes(xlValue).MinimumScale = yMin
31                End If
32                If chOb.Chart.Axes(xlCategory).MaximumScale <> ThisxMax Then
33                    chOb.Chart.Axes(xlCategory).MaximumScale = ThisxMax
34                End If
35                If chOb.Chart.Axes(xlCategory).MinimumScale <> ThisxMin Then
36                    chOb.Chart.Axes(xlCategory).MinimumScale = ThisxMin
37                End If
38            Next chOb
39        Next ws
40        Exit Sub
ErrHandler:
41        Throw "#RescaleAllChartsInBook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


