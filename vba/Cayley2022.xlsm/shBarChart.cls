VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shBarChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Calculate
' Author    : Philip Swannell
' Date      : 27-Sep-2016
' Purpose   : Make only the banks where the "Include in Chart?" flag is TRUE appear in the chart
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Calculate()
1         On Error GoTo ErrHandler
2         If Me.ChartObjects.Count > 0 Then
3             If ActiveSheet Is Me Then
4                 FixBarChart False
5             End If
6         End If
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#" & Me.Name & ".Worksheet_Calculate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 30-Nov-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
1         On Error GoTo ErrHandler
2         If Not Application.Intersect(Target, RangeFromSheet(Me, "SortBy")) Is Nothing Then
3             Me.Calculate
4         End If

5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
