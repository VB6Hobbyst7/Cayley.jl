VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public g_LastChangeTime As Double

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim DataToFilter As Variant
          Dim NumTrades
          Dim RegKey As String
          Dim Res
          Dim TradeAttribute As String
          Dim TradeRange As Range
          Dim TradesRange As Range
          Dim TradeType As String

1         On Error GoTo ErrHandler
2         If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
3             Cancel = True
4             RegKey = RangeFromSheet(shHiddenSheet, "RegKeys").Cells(1, Target.Column - RangeFromSheet(Me, "TheFilters").Column + 1).Value
5             DataToFilter = Application.Intersect(getTradesRange(NumTrades), Target.EntireColumn).Value
6             If NumTrades = 0 Then DataToFilter = Empty
7             Res = ShowRegularExpressionDialog(Target, RegKey, DataToFilter, , "Filter Trades", "Show trades where:", True, RegKey, False)
8             Exit Sub
9         End If
          'Calling GetTradesRange clears out the clipboard, but we may need the clipboard for entering amortising notionals, so reload the clipboard.
10        If Application.CutCopyMode = xlCopy Then
              Dim CopiedRange As Range
11            GetCopiedRange CopiedRange
12        End If
13        Set TradesRange = getTradesRange(NumTrades)
14        If Not CopiedRange Is Nothing Then
15            CopiedRange.Copy
16        End If

          'Double click in row beneath last trade
17        If Not Application.Intersect(Target, TradesRange.Rows(IIf(NumTrades = 0, 1, TradesRange.Rows.Count + 1))) Is Nothing Then
18            Cancel = True
19            ShowMenu "AddOneTrade"
20            Exit Sub
21        End If

22        If NumTrades > 0 Then
23            If Not Application.Intersect(Target, TradesRange) Is Nothing Then
24                Set TradeRange = Application.Intersect(Target.EntireRow, TradesRange)

25                TradeType = TradeRange.Cells(1, gCN_TradeType).Value
26                TradeAttribute = Application.Intersect(Target.EntireColumn, TradesRange.Rows(0)).Value
27                If TradeAttribute = "" Then    ' madness of two rows of headers with merged cells...
28                    TradeAttribute = Application.Intersect(Target.EntireColumn, TradesRange.Rows(-1)).Value
29                End If
30                TradeDoubleClickHandler Target, TradeRange, TradeType, TradeAttribute, Cancel
31            End If
32        End If

33        If Not Cancel Then RunButtonAtActiveCell

34        Exit Sub
ErrHandler:
35        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
          Dim NumTrades
          Dim TradesRange As Range
1         On Error GoTo ErrHandler
2         Set TradesRange = getTradesRange(NumTrades)
3         If NumTrades > 0 Then
4             If Not Application.Intersect(Target, TradesRange) Is Nothing Then
5                 Cancel = True
6                 ShowMenu "FromRightClick"
7                 Exit Sub
8             End If
9         End If
10        If Target.Areas(1).Cells(1, 1).Address = TradesRange.Cells(TradesRange.Rows.Count + IIf(NumTrades = 0, 0, 1), gCN_TradeType).Address Then
11            Cancel = True
12            ShowMenu "AddOneTrade"
13            Exit Sub
14        End If
15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#Worksheet_BeforeRightClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Worksheet_Calculate
' Author    : Philip Swannell
' Date      : 13-Apr-2016
' Purpose   : Calculating the Portfolio refreshes the numbers on it!
'---------------------------------------------------------------------------------------
Private Sub Worksheet_Calculate()
          Static InEvent As Boolean
          Dim CopyOfErr As String
1         On Error GoTo ErrHandler
2         If InEvent Then Exit Sub
3         If gBlockCalculateEvent Then Exit Sub
4         If ActiveSheet Is Nothing Then Exit Sub
5         If Not ActiveSheet Is Me Then Exit Sub
6         InEvent = True
7         If IsShiftKeyDown Then
8             XVAFrontEndMain True, False, False, False, False, False, True, False
9         End If
10        InEvent = False
11        Exit Sub
ErrHandler:
12        CopyOfErr = "#shPortfolio.Worksheet_Calculate (line " & CStr(Erl) + "): " & Err.Description & "!"
13        InEvent = False
14        SomethingWentWrong CopyOfErr, vbCritical
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
          Dim a As Range
          Dim CopyOfErr As String
          Dim i As Long
          Dim SUH As SolumAddin.clsScreenUpdateHandler

1         On Error GoTo ErrHandler

2         If InEvent Then Exit Sub
3         If gBlockChangeEvent Then Exit Sub
4         InEvent = True
5         g_LastChangeTime = Now()
6         For Each a In Target.Areas
7             For i = 1 To a.Rows.Count
8                 If a.Rows(i).Hidden Then
9                     Application.Undo
10                    Throw "That selection includes hidden trades, because of the filters at the top of the sheet. Trades cannot be edited whilst hidden.", True
11                End If
12            Next i
13        Next a

14        Set SUH = CreateScreenUpdateHandler()

15        If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
16            FilterTradesRange
17            SetTradesRangeColumnWidths
18        Else
19            FormatTradesRange Target
20            FilterTradesRange
21            SetTradesRangeColumnWidths
22        End If

23        InEvent = False
24        Exit Sub
ErrHandler:
25        CopyOfErr = "#shPortfolio.Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
26        InEvent = False
27        SomethingWentWrong CopyOfErr
End Sub
