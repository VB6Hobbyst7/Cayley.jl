VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shCounterpartyViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
1         On Error GoTo ErrHandler
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim TheTradeID As Variant

2         Set SPH = CreateSheetProtectionHandler(Me)
3         If Target.Address = RangeFromSheet(Me, "Inc_Hypotheticals?").Address Then
4             Cancel = True
5             Target.Value = Not sEquals(Target.Value, True)
6         ElseIf Target.Address = RangeFromSheet(Me, "SelectedCpty").Address Then
              Dim Allowed As Variant
              Dim Chosen As Variant
7             Cancel = True
8             Allowed = CounterpartiesFromJulia()
9             Chosen = ShowSingleChoiceDialog(Allowed, , , , , "Select Counterparty", , Target.Offset(0, 1), , "Counterparty")
10            If Not IsEmpty(Chosen) Then
11                Target.Value = Chosen
12            End If
13        ElseIf Not Application.Intersect(Target, RangeFromSheet(Me, "TradeDataWithHeaders")) Is Nothing Then
14            If Target.Row > RangeFromSheet(Me, "TradeDataWithHeaders").Row Then
15                TheTradeID = Me.Cells(Target.Row, RangeFromSheet(Me, "TradeDataWithHeaders").Column)
16                If VarType(TheTradeID) = vbString Then
17                    If Len(TheTradeID) > 0 Then
18                        Cancel = True
19                        Application.ScreenUpdating = False
20                        shTradeViewer.Activate
21                        Application.GoTo shTradeViewer.Cells(1, 1)
22                        Application.GoTo RangeFromSheet(shTradeViewer, "SelectedTrade")
23                        RangeFromSheet(shTradeViewer, "SelectedTrade").Value = TheTradeID
24                    End If
25                End If
26            End If
27        End If
28        Exit Sub
ErrHandler:
29        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!", , MsgBoxTitle()
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 03-Nov-2015
' Purpose   : Make response from the SelectedCpty drop down be instant
'---------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
1         On Error GoTo ErrHandler
2         If Not gBlockChangeEvent Then
3             If Not Application.Intersect(Target, Application.Union(RangeFromSheet(Me, "SelectedCpty"), RangeFromSheet(Me, "Inc_Hypotheticals?"))) Is Nothing Then
4                 UpdateCounterpartyViewerSheet
5             End If
6         End If
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

