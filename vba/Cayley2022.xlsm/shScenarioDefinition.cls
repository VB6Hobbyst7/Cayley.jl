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
          Dim EnableFlags
          Dim HH As Long
          Dim Res

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         EnableFlags = createmissing()

4         Select Case Target.Address
              Case RangeFromSheet(Me, "UseSpeedGrid").Address
5                 CurrentChoice = UCase(CStr(Target.Value))
6                 If RangeFromSheet(Me, "ShocksDerivedFrom").Value = "Custom" Then
7                     Alternatives = sArrayStack("TRUE First set 'ShocksDerivedFrom' to 'History', not 'Custom'", "FALSE")
8                     EnableFlags = sArrayStack(False, True)
9                 Else
10                    Alternatives = sArrayStack("TRUE", "FALSE")
11                End If

12            Case RangeFromSheet(Me, "VaryGridWidth").Address
13                Alternatives = sArrayStack("TRUE", "FALSE")
14                CurrentChoice = UCase(CStr(Target.Value))

15            Case RangeFromSheet(Me, "ShocksDerivedFrom").Address
16                Alternatives = sArrayStack("History", "Custom")
17                CurrentChoice = Target.Value
18            Case RangeFromSheet(Me, "HistoryStart").Address
19                Alternatives = sArrayStack("01-Jan-2000", "01-Jan-2007", "01-Jan-2009", "15-Apr-2013")
20                CurrentChoice = Format(Target.Value2, "dd-mmm-yyyy")
21            Case RangeFromSheet(Me, "AllocationByYear").Address
22                HH = GetHedgeHorizon()
                  Dim Extra As String
23                If HH > 5 Then Extra = sConcatenateStrings(sReshape(":0", HH - 5, 1), "")
24                Alternatives = sArrayStack("0:0:1:1:1" & Extra, "0:0:3:1:1" & Extra, "1:1:1:1:1" & Extra)
25                CurrentChoice = Target.Value
26        End Select

27        If Not IsEmpty(Alternatives) Then
28            Cancel = True
29            If sNRows(Alternatives) < 20 Then
30                Res = ShowCommandBarPopup(Alternatives, , EnableFlags, CurrentChoice)
31                If Res = "#Cancel!" Then Exit Sub
32                Target.Value = Res
33                Application.GoTo Target
34            Else
35                Alternatives = sArrayMakeText(Alternatives)
36                Res = ShowSingleChoiceDialog(Alternatives, , , , , , , Target.offset(0, 1), , "Scenario Definition")
37                If Not IsEmpty(Res) Then
38                    Target.Value = Res
39                End If
40            End If
41        End If

42        Exit Sub
ErrHandler:
43        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!"
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

