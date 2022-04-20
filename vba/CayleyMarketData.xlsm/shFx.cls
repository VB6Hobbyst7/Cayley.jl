VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 30 March 2017
' Purpose   : Filtering...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static Inevent As Boolean
          Dim Message As String
          Dim RegKeys As Variant
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If Inevent Then Exit Sub

3         Inevent = True

4         If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
5             RegKeys = sArrayConcatenate("SMDFx", sArrayTranspose(Me.Range("TheFilters").Offset(2)))
6             FilterRangeByHidingRows RangeFromSheet(Me, "TheFilters"), TheRangeToFilter(), , Message, , RegKeys
7             Message = Message
8             With RangeFromSheet(Me, "TheFilters").Cells(0, 1)
9                 If Message <> .Value Then
10                    Set SPH = CreateSheetProtectionHandler(Me)
11                    .Value = Message
12                    .Font.Color = 8421504
13                End If
14            End With
15        End If
16        Inevent = False
17        Exit Sub
ErrHandler:
18        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
19        Inevent = False
End Sub

Private Function TheRangeToFilter() As Range
1         On Error GoTo ErrHandler
2         Set TheRangeToFilter = sExpandDown(Me.Range("TheFilters").Rows(4))
3         Exit Function
ErrHandler:
4         Throw "#TheRangeToFilter (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim AttributeName As String
          Dim DataToFilter As Range
          Dim RegKey As String
          Dim Res
1         On Error GoTo ErrHandler

2         If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
3             Cancel = True
4             AttributeName = CStr(Target.Offset(2).Value)
5             RegKey = "SMDFx" & AttributeName
6             Set DataToFilter = Application.Intersect(TheRangeToFilter(), Target.EntireColumn)
7             Res = ShowRegularexpressionDialog(Target, AttributeName, DataToFilter, , "Filter rows", "Show rows where:", True, RegKey, False)
8             Exit Sub
9         End If

10        Exit Sub
ErrHandler:
11        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

