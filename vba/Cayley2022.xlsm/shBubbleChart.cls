VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shBubbleChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Calculate
' Author    : Philip Swannell
' Date      : 16-Nov-2016
' Purpose   : Can't figure out how to hook the title of the bubble chart to a cell's value,
'             despite this being easy on (say) the bar chart
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Calculate()
          Dim SPH As clsSheetProtectionHandler
1         On Error GoTo ErrHandler
          Dim ch As ChartObject

2         If Me.ChartObjects.Count > 0 Then
3             Set ch = Me.ChartObjects(1)
4             Set SPH = CreateSheetProtectionHandler(Me)
5             If ch.Chart.ChartTitle.Caption <> Me.Range("ChartTitle").Value Then
6                 ch.Chart.ChartTitle.Caption = Me.Range("ChartTitle").Value
7             End If
8             ch.Chart.Axes(xlCategory).MinimumScale = RangeFromSheet(Me, "FxBreakEvenFloor").Value
9         End If

10        Exit Sub
ErrHandler:
11        SomethingWentWrong "#Worksheet_Calculate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
1         On Error GoTo ErrHandler
2         If Target.Address = RangeFromSheet(Me, "FxBreakEvenFloor").Address Then
3             Me.Calculate
4         End If

5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
