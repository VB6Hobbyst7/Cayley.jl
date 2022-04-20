VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shTradesViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_BeforeDoubleClick
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Filtering dialogs
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim AttributeName As String
          Dim DataToFilter As Range
          Dim RegKey As String
          Dim Res

          'If no trades are displayed then the range TheData will not exist
1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         If IsInCollection(Me.Names, "TheData") Then
4             If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
5                 Cancel = True
6                 AttributeName = CStr(Target.offset(2).Value)
7                 RegKey = "Cayley" & AttributeName
8                 Set DataToFilter = Application.Intersect(RangeFromSheet(Me, "TheData"), Target.EntireColumn)
9                 Res = ShowRegularExpressionDialog(Target, AttributeName, DataToFilter, , _
                      "Filter Trades", "Show trades where:", True, RegKey, False)
10                Exit Sub
11            End If
12        End If

13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Filtering...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
          Dim Message As String
          Dim RegKeys
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If InEvent Then Exit Sub
3         If gBlockChangeEvent Then Exit Sub
4         InEvent = True

          'If no trades are displayed then the range TheData will not exist
5         If IsInCollection(Me.Names, "TheData") Then
6             If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
7                 RegKeys = sArrayConcatenate("Cayley", sArrayTranspose(Me.Range("TheHeaders")))
8                 FilterRangeByHidingRows RangeFromSheet(Me, "TheFilters"), RangeFromSheet(Me, "TheData"), _
                      "trade", Message, , RegKeys
9                 If Message <> RangeFromSheet(Me, "NumTrades").Value Then
10                    Set SPH = CreateSheetProtectionHandler(Me)
11                    RangeFromSheet(Me, "NumTrades").Value = Message
12                End If
13            End If
14        End If
15        InEvent = False
16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
18        InEvent = False
End Sub


