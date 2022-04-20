VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Activate
' Author    : Philip Swannell
' Date      : 17-Jan-2017
' Purpose   : Have not been able to figure out why the menu button on the Table sheet
'             sometimes goes AWOL, so realign the menu buttons (on all sheets
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
1         On Error Resume Next
2         AlignMenuButtons
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_BeforeDoubleClick
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Filtering dialogs
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim AttributeName As String
          Dim ClickableRange As Range
          Dim DataToFilter As Range
          Dim RegKey As String
          Dim Res
1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         With RangeFromSheet(Me, "TheTable")
4             Set ClickableRange = .offset(1).Resize(.Rows.Count - 1)
5         End With

6         If Not Application.Intersect(Target, ClickableRange) Is Nothing Then
7             Cancel = True
8             PrepareForCalculation Application.Intersect(Target.EntireRow, _
                  ClickableRange.Columns(1)).Value, True, False, False
9             shCreditUsage.Activate
10            JuliaLaunchForCayley
11            RunCreditUsageSheet "Standard", True, False, True
12        ElseIf Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
13            Cancel = True
14            AttributeName = Replace(CStr(Target.offset(2).Value), vbLf, " ")
15            RegKey = "Cayley" & AttributeName
16            Set DataToFilter = Application.Intersect(RangeFromSheet(Me, "TheTableNoHeaders"), Target.EntireColumn)
17            Res = ShowRegularExpressionDialog(Target, AttributeName, DataToFilter, , _
                  "Filter rows", "Show rows where:", True, RegKey, False)
18            Exit Sub
19        End If

20        Exit Sub
ErrHandler:
21        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Filtering...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
          Dim i As Long
          Dim Message As String
          Dim RegKeys As Variant
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If InEvent Then Exit Sub
3         If gBlockChangeEvent Then Exit Sub
4         InEvent = True

5         If Not Application.Intersect(Target, RangeFromSheet(Me, "TheFilters")) Is Nothing Then
6             RegKeys = sArrayConcatenate("Cayley", sArrayTranspose(Me.Range("TheHeaders")))
7             For i = 1 To sNRows(RegKeys)
8                 RegKeys(i, 1) = Replace(RegKeys(i, 1), vbLf, " ")
9             Next
10            FilterRangeByHidingRows RangeFromSheet(Me, "TheFilters"), _
                  RangeFromSheet(Me, "TheTableNoHeaders"), , Message, , RegKeys
11            Message = Message & "    Double-click to drill down."
12            If Message <> RangeFromSheet(Me, "Message").Value Then
13                Set SPH = CreateSheetProtectionHandler(Me)
14                With RangeFromSheet(Me, "Message")
15                    .Value = Message
16                    .Font.Color = g_Col_GreyText
17                End With
18            End If
19        End If
20        InEvent = False
21        Exit Sub
ErrHandler:
22        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
23        InEvent = False
End Sub


