VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shScenarioDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim Alternatives
          Dim CurrentChoice
          Dim HH As Long
          Dim Res

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         Select Case Target.Address

              Case RangeFromSheet(Me, "UseSpeedGrid").Address, _
                  RangeFromSheet(Me, "VaryGridWidth").Address
4                 Alternatives = sArrayStack("TRUE", "FALSE")
5                 CurrentChoice = UCase(CStr(Target.Value))

6             Case RangeFromSheet(Me, "ShocksDerivedFrom").Address
7                 Alternatives = sArrayStack("History", "Custom")
8                 CurrentChoice = Target.Value
9             Case RangeFromSheet(Me, "HistoryStart").Address
10                Alternatives = sArrayStack("01-Jan-2000", "01-Jan-2007", "01-Jan-2009", "15-Apr-2013")
11                CurrentChoice = Format(Target.Value2, "dd-mmm-yyyy")
12            Case RangeFromSheet(Me, "AllocationByYear").Address
13                HH = GetHedgeHorizon()
                  Dim Extra As String
14                If HH > 5 Then Extra = sConcatenateStrings(sReshape(":0", HH - 5, 1), "")
15                Alternatives = sArrayStack("0:0:1:1:1" & Extra, "0:0:3:1:1" & Extra, "1:1:1:1:1" & Extra)
16                CurrentChoice = Target.Value
17        End Select

18        If Not IsEmpty(Alternatives) Then
19            Cancel = True
20            If sNRows(Alternatives) < 20 Then
21                Res = ShowCommandBarPopup(Alternatives, , , CurrentChoice, Target)
22                If Res = "#Cancel!" Then Exit Sub
23                Target.Value = Res
24                Application.GoTo Target
25            Else
26                Alternatives = sArrayMakeText(Alternatives)
27                Res = ShowSingleChoiceDialog(Alternatives, , , , , , , Target.offset(0, 1), , "Scenario Definition")
28                If Not IsEmpty(Res) Then
29                    Target.Value = Res
30                End If
31            End If
32        End If

33        Exit Sub
ErrHandler:
34        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!"

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
          Dim CopyOfErr As String
          Dim TriggerCells As Range

1         On Error GoTo ErrHandler
2         If gBlockChangeEvent Then Exit Sub
3         If InEvent Then Exit Sub
4         InEvent = True

5         Set TriggerCells = Application.Union(RangeFromSheet(Me, "HistoryStart"), _
              RangeFromSheet(Me, "ShocksDerivedFrom"), _
              RangeFromSheet(Me, "UseSpeedGrid"), _
              RangeFromSheet(Me, "SpeedGridWidth"), _
              RangeFromSheet(Me, "HighFxSpeed"), _
              RangeFromSheet(Me, "LowFxSpeed"), _
              RangeFromSheet(Me, "VaryGridWidth"), _
              RangeFromSheet(Me, "SpeedGridBaseVol"), _
              RangeFromSheet(Me, "AnnualReplenishment"))

6         If Not Application.Intersect(Target, TriggerCells) Is Nothing Then
7             RefreshScenarioDefinition False
8         End If

9         InEvent = False
10        Exit Sub
ErrHandler:
11        CopyOfErr = "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
12        InEvent = False
13        SomethingWentWrong CopyOfErr
End Sub


