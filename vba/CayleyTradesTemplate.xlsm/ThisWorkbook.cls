VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
          'The trades workbook is not the store of data (the csv files are) so we don't want to be asked if we want to save it.
1         ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

          Dim AttributeName As String
          Dim DataRange As Range
          Dim DataToFilter As Range
          Dim FilterRange As Range
          Dim HeaderRange As Range
          Dim Headers As Variant
          Dim lo As ListObject
          Dim RegKey As String
          Dim RowDescriptor

1         On Error GoTo ErrHandler

2         Set lo = Sh.ListObjects(1)
3         Set HeaderRange = lo.HeaderRowRange
4         Set FilterRange = RangeFromSheet(Sh, "TheFilters")
5         Set DataRange = lo.DataBodyRange

6         Headers = sArrayTranspose(HeaderRange.Value)

7         If Not Application.Intersect(Target, FilterRange) Is Nothing Then
8             RowDescriptor = IIf(Sh.Name = "Amortisation", "row", "trade")

9             Cancel = True
10            AttributeName = CStr(Target.Offset(2).Value)
11            RegKey = "CayleyTrades" & AttributeName
12            Set DataToFilter = Application.Intersect(DataRange, Target.EntireColumn)
13            ShowRegularExpressionDialog Target, AttributeName, DataToFilter, , "Filter " & RowDescriptor & "s", "Show " & RowDescriptor & "s where:", True, RegKey, False, DataRange
14        End If

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#Workbook_SheetBeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

          Static InEvent As Boolean
          
          Dim DataRange As Range
          Dim FilterRange As Range
          Dim HeaderRange As Range
          Dim lo As ListObject
          Dim Message As String
          Dim MessageRange
          Dim RegKeys
          Dim RowDescriptor As String
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If InEvent Then Exit Sub
3         If Sh.ListObjects.Count >= 1 Then
4             InEvent = True
5             RowDescriptor = IIf(Sh.Name = "Amortisation", "row", "trade")
              
7             Set lo = Sh.ListObjects(1)
8             Set HeaderRange = lo.HeaderRowRange

10            Set FilterRange = RangeFromSheet(Sh, "TheFilters")
11            Set MessageRange = FilterRange.Cells(0, 1)
12            Set DataRange = lo.DataBodyRange
13            If Not Application.Intersect(Target, FilterRange) Is Nothing Then
14                RegKeys = sArrayConcatenate("CayleyTrades", sArrayTranspose(HeaderRange))
15                FilterRangeByHidingRows FilterRange, DataRange, RowDescriptor, Message, True, RegKeys
16                If Message <> MessageRange.Value Then
17                    Set SPH = CreateSheetProtectionHandler(Sh)
18                    MessageRange.Value = Message
19                End If
20            End If
21            InEvent = False
22        End If
23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
25        InEvent = False
End Sub

