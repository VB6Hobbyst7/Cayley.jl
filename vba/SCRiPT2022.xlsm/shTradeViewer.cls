VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shTradeViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim MatchID
          Dim TradeData
          Dim TradeDescriptions
          Dim TradeIDs
          Dim Trades As Dictionary
1         On Error GoTo ErrHandler
2         If Target.Address = RangeFromSheet(Me, "SelectedTrade").Address Then
              Dim Chosen As Variant
3             Cancel = True
4             If IsPFEDataAvailable("trade", True) Then
5                 Set Trades = gResults("Trades")
6                 TradeData = HStack(Trades("TradeID"), Trades("ValuationFunction"), Trades("Counterparty"))
7                 TradeData = sMChoose(TradeData, sArrayNot(sArrayEquals(gWHATIF, sSubArray(TradeData, 1, 3, , 1))))
8                 TradeData = sSortedArray(TradeData)
9                 TradeDescriptions = sJustifyArrayOfStrings(TradeData, "Tahoma", 8, vbTab)
10                TradeIDs = sSubArray(TradeData, 1, 1, , 1)
11                Chosen = ShowSingleChoiceDialog(TradeDescriptions, , , , , "Select Trade", , Target.Offset(0, 1), , "TradeID")
12                If Not IsEmpty(Chosen) Then
13                    MatchID = sMatch(Chosen, TradeDescriptions)
14                    Chosen = TradeIDs(MatchID, 1)
15                    Target.Value = "'" + Chosen
16                End If
17            End If
18        End If
19        Exit Sub
ErrHandler:
20        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!", , MsgBoxTitle()
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
3             If Not Application.Intersect(Target, RangeFromSheet(Me, "SelectedTrade")) Is Nothing Then
4                 UpdateTradeViewerSheet
5             End If
6         End If
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
