VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shExportToTMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim Res As Variant
          Dim RngBooleans As Range
          Dim RngPasteTo As Range
          Dim RngScenarioDefinitions
          Dim SPH As clsSheetProtectionHandler
          Const RegKeyWhereToExport = "CayleyWhereToExport"
          Dim i As Long
          
1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         Set SPH = CreateSheetProtectionHandler(Me)

4         Cancel = True

5         Set RngBooleans = Application.Union(RangeFromSheet(Me, "Scenarios").Columns(1), _
              RangeFromSheet(Me, "FeedRates"), _
              RangeFromSheet(Me, "ExportTrades"), _
              RangeFromSheet(Me, "ExportMarketData"), _
              RangeFromSheet(Me, "ExportTable"), _
              RangeFromSheet(Me, "ExportCharts"))
              
6         Set RngScenarioDefinitions = RangeFromSheet(Me, "Scenarios").Columns(2)

7         If Not Application.Intersect(Target, RngBooleans) Is Nothing Then
8             Target.Value = Not sEquals(Target.Value, True)
9         ElseIf Not Application.Intersect(Target, RngScenarioDefinitions) Is Nothing Then
10            If Right(Target.Value, 4) = ".sdf" Then
11                If sFileExists(Target.Value) Then
12                    AddFileToMRU gRegKey_Defn, CStr(Target.Value)
13                End If
14            End If

15            Res = GetOpenFilenameWrap(gRegKey_Defn, "Scenario Definition Files (*.SDF),*.SFD", , "Select Scenario Definition File(s)", , True, True, Target)
16            Select Case VarType(Res)
                  Case vbBoolean
17                    Exit Sub
18                Case vbString
19                    Target.Value = "'" & Res
20                    Target.offset(0, -1).Value = True
21                Case Is > vbArray
22                    Res = sArrayTranspose(Res)
23                    For i = 1 To sNRows(Res)
24                        AddFileToMRU gRegKey_Defn, CStr(Res(i, 1))
25                    Next
                      
26                    Set RngPasteTo = Target.Resize(sNRows(Res))
27                    Set RngPasteTo = Application.Intersect(RngPasteTo, RangeFromSheet(Me, "Scenarios"))
28                    Application.GoTo RngPasteTo
29                    RngPasteTo.Value = sArrayExcelString(Res)
30                    RngPasteTo.offset(0, -1).Value = True
31            End Select
32        ElseIf Target.Address = RangeFromSheet(Me, "WhereToExport").Address Then
33            If sFolderExists(CStr(Target.Value)) Then
34                AddFileToMRU RegKeyWhereToExport, CStr(Target.Value)
35            End If
36            Res = FolderPicker(Target.Value, , "Select Where To Export", RegKeyWhereToExport, True, Target)
37            If VarType(Res) <> vbString Then Exit Sub
38            Target.Value = "'" & Res
39        Else
40            Cancel = False
41        End If

42        Exit Sub
ErrHandler:
43        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & _
              Err.Description & "!", vbExclamation, "Worksheet_BeforeDoubleClick"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

1         SaveETMSToRegistry

End Sub

