VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim AttributeName As String
          Dim BankName As String
          Dim ClickRangeFxNW As Range
          Dim ClickRangeRNW As Range
          Dim ColNoCP As Variant
          Dim ColNoFxNW
          Dim ColNoRNW As Variant
          Dim DataRange As Range
          Dim DataToFilter As Range
          Dim FilterRange As Range
          Dim HeaderRange As Range
          Dim Headers As Variant
          Dim lo As ListObject
          Dim Prompt As String
          Dim RegKey As String

1         On Error GoTo ErrHandler

2         Set lo = Me.ListObjects(1)
3         Set lo = Me.ListObjects(1)
4         Set HeaderRange = lo.HeaderRowRange
5         Set FilterRange = lo.HeaderRowRange.Offset(-2)
6         Set DataRange = lo.DataBodyRange

7         Headers = sArrayTranspose(HeaderRange.Value)
8         ColNoRNW = sMatch("Rates Notional Weights", Headers)
9         If Not IsNumber(ColNoRNW) Then Throw "Cannot find column headed 'Rates Notional Weights' on sheet Summary"
10        ColNoFxNW = sMatch("Fx Notional Weights", Headers)
11        If Not IsNumber(ColNoFxNW) Then Throw "Cannot find column headed 'Fx Notional Weights' on sheet Summary"

12        ColNoCP = sMatch("CPTY_PARENT", Headers)
13        If Not IsNumber(ColNoRNW) Then Throw "Cannot find column headed 'CPTY_PARENT' on sheet Summary"
14        Set ClickRangeRNW = DataRange.Columns(ColNoRNW)
15        Set ClickRangeFxNW = DataRange.Columns(ColNoFxNW)
16        If Not Application.Intersect(ClickRangeRNW, Target) Is Nothing Then
17            Cancel = True
18            BankName = Application.Intersect(Me.ListObjects(1).DataBodyRange.Columns(ColNoCP), Target.EntireRow).Value
19            Prompt = "Edit Rates Notional Weights for " + BankName + "?"
20            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
21            EditNotionalWeights BankName, Target.Value, True
22        ElseIf Not Application.Intersect(ClickRangeFxNW, Target) Is Nothing Then
23            Cancel = True
24            BankName = Application.Intersect(Me.ListObjects(1).DataBodyRange.Columns(ColNoCP), Target.EntireRow).Value
25            Prompt = "Edit Fx Notional Weights for " + BankName + "?"
26            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
27            EditNotionalWeights BankName, Target.Value, False
28        ElseIf Not Application.Intersect(Target, FilterRange) Is Nothing Then
29            Cancel = True
30            AttributeName = CStr(Target.Offset(2).Value)
31            RegKey = "CayleyLines" & AttributeName
32            Set DataToFilter = Application.Intersect(DataRange, Target.EntireColumn)
33            ShowRegularExpressionDialog Target, AttributeName, DataToFilter, , "Filter Rows", "Show rows where:", True, RegKey, False
34        End If

35        Exit Sub
ErrHandler:
36        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 08-Oct-2016
' Purpose   : Filtering...
'---------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
          Dim DataRange As Range
          Dim FilterRange As Range
          Dim HeaderRange As Range
          Dim lo As ListObject
          Dim Message As String
          Dim MessageRange
          Dim RegKeys
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If InEvent Then Exit Sub
3         If Me.ListObjects.Count >= 1 Then
4             InEvent = True
5             Set lo = Me.ListObjects(1)
6             Set HeaderRange = lo.HeaderRowRange
7             Set FilterRange = lo.HeaderRowRange.Offset(-2)
8             Set MessageRange = FilterRange.Cells(0, 1)
9             Set DataRange = lo.DataBodyRange
10            If Not Application.Intersect(Target, FilterRange) Is Nothing Then
11                RegKeys = sArrayConcatenate("CayleyLines", sArrayTranspose(HeaderRange))
12                FilterRangeByHidingRows FilterRange, DataRange, "row", Message, , RegKeys
13                If Message <> MessageRange.Value Then
14                    Set SPH = CreateSheetProtectionHandler(Me)
15                    MessageRange.Value = Message
16                End If
17            End If
18            InEvent = False
19        End If
20        Exit Sub
ErrHandler:
21        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
22        InEvent = False
End Sub

