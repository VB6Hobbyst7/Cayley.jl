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

2         If Me.ChartObjects.Count > 0 Then
3             If Me.ChartObjects(1).Chart.ChartTitle.Caption <> Me.Range("ChartTitle").Value Then
4                 Set SPH = CreateSheetProtectionHandler(Me)
5                 Me.ChartObjects(1).Chart.ChartTitle.Caption = Me.Range("ChartTitle").Value
6             End If
7         End If

8         Exit Sub
ErrHandler:
9         Throw "#Worksheet_Calculate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

