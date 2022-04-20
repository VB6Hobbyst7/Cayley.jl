Attribute VB_Name = "modUndo"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modUndo
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Code to make it easy to implement Application.OnUndo
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Public Const MAX_CELLS_FOR_UNDO = 1000000
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BackUpRange
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : General purpose routine to backup a range of cells (values, formulas, formats)
'             to a hidden sheet of this workbook. Companion function RestoreRange can be hooked
'             up to Ctrl+Z via Application.OnUndo.
' -----------------------------------------------------------------------------------------------------------------------
Sub BackUpRange(RangeToBackup As Range, SheetToBackUpTo As Worksheet, Optional RangeToSelectAfterUndo As Range, Optional ByVal ValuesOnly As Boolean)
          Dim c As Range
          Dim SUH As clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler()
3         Set XSH = CreateExcelStateHandler(, , False)    ' For speed, we don't want events firing

4         If IsUndoAvailable(SheetToBackUpTo) Then
5             CleanUpUndoBuffer SheetToBackUpTo
6         End If

7         SheetToBackUpTo.Names("xxx_UndoBufferIsEmpty").Delete

8         If RangeToBackup.Cells.CountLarge > MAX_CELLS_FOR_UNDO Then
9             CleanUpUndoBuffer SheetToBackUpTo
10            Exit Sub
11        End If

12        If RangeToSelectAfterUndo Is Nothing Then
13            Set RangeToSelectAfterUndo = RangeToBackup
14        End If

15        If ValuesOnly Then
16            For Each c In RangeToBackup.Areas
17                If Not IsFalse(c.HasFormula) Then
18                    ValuesOnly = False
19                    Exit For
20                End If
21            Next c
22        End If

23        SheetToBackUpTo.Names.Add "xxx_BackedUpBookName", RangeToBackup.Parent.Parent.Name
24        SheetToBackUpTo.Names.Add "xxx_BackedUpSheetName", RangeToBackup.Parent.Name
25        SheetToBackUpTo.Names.Add "xxx_BackedUpRangeAddress", RangeToBackup.address
26        SheetToBackUpTo.Names.Add "xxx_VisibleRangeAddress", ActiveWindow.VisibleRange.address
27        SheetToBackUpTo.Names.Add "xxx_ValuesOnly", ValuesOnly

28        If Not RangeToSelectAfterUndo Is Nothing Then
29            SheetToBackUpTo.Names.Add "xxx_RangeToSelectAfterUndoAddress", RangeToSelectAfterUndo.address
30        End If

31        If ValuesOnly Then
32            For Each c In RangeToBackup.Areas
                  'avoid copy and paste which is slow...
33                Set TargetRange = SheetToBackUpTo.Range(c.address)
34                MyPaste TargetRange, c.Value2
35            Next c
36        Else
37            For Each c In RangeToBackup.Areas
38                Set TargetRange = SheetToBackUpTo.Range(c.address)
39                c.Copy
40                Application.DisplayAlerts = False
41                TargetRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
42            Next c
43        End If

44        Application.CutCopyMode = False

45        Exit Sub
ErrHandler:
46        Throw "#BackUpRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function IsFalse(x) As Boolean
1         If VarType(x) = vbBoolean Then
2             IsFalse = Not (x)
3         End If
End Function

Sub CleanUpUndoBuffer(SheetToClean As Worksheet)
          Dim N As Name
          Dim S As Shape
          Dim x As Long

1         On Error GoTo ErrHandler
2         For Each N In SheetToClean.Names
3             N.Delete
4         Next
5         For Each S In SheetToClean.Shapes
6             S.Delete
7         Next S

8         SheetToClean.UsedRange.EntireRow.Delete
          'Reset the UsedRange
9         x = SheetToClean.UsedRange.Rows.Count

10        For Each N In ThisWorkbook.Names
11            If InStr(N.RefersTo, "#REF!") > 0 Then
12                N.Delete
13            End If
14        Next
15        SheetToClean.Names.Add "xxx_UndoBufferIsEmpty", "UndoBufferIsEmpty"
16        Exit Sub
ErrHandler:
17        Throw "#CleanUpUndoBuffer (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub RestoreRangeTwice()
1         On Error GoTo ErrHandler
2         If IsUndoAvailable(shUndo) And IsUndoAvailable(shUndo2) Then
3             RestoreRangeCore shUndo
4             RestoreRangeCore shUndo2
5         Else
6             MsgBoxPlus "Undo is not available", vbExclamation, "Undo"
7         End If
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#RestoreRangeTwice (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, "Restore Range Twice"
End Sub

Sub RestoreRangeFromUndoBuffer2()
1         On Error GoTo ErrHandler

2         If IsUndoAvailable(shUndo2) Then
3             RestoreRangeCore shUndo2
4         Else
5             MsgBoxPlus "Undo is not available", vbExclamation, "Undo"
6         End If

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#RestoreRangeFromUndoBuffer2 (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

Sub RestoreRange()
1         On Error GoTo ErrHandler

2         If IsUndoAvailable(shUndo) Then
3             RestoreRangeCore shUndo
4         Else
5             MsgBoxPlus "Undo is not available", vbExclamation, "Undo"
6         End If

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#RestoreRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RestoreRange
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : See comments for BackUpRange
' -----------------------------------------------------------------------------------------------------------------------
Sub RestoreRangeCore(SheetToRestoreFrom As Worksheet)
          Dim address As Variant
          Dim AreaAddresses As Variant
          Dim BackedUpBookName As String
          Dim BackedUpRangeAddress As String
          Dim BackedUpSheetName As String
          Dim OrigSelection As Range
          Dim origViewPort As Range
          Dim RangeToSelectAfterUndoAddress As String
          Dim SourceRange
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TargetRange As Range
          Dim ValuesOnly As Boolean
          Dim VisibleRangeAddress As String
          Dim wbTarget As Excel.Workbook
          Dim wsTarget As Worksheet
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler()

3         If Not IsUndoAvailable(SheetToRestoreFrom) Then Exit Sub

4         If Not IsInCollection(SheetToRestoreFrom.Names, "xxx_BackedUpBookName") Or _
              Not IsInCollection(SheetToRestoreFrom.Names, "xxx_BackedUpSheetName") Or _
              Not IsInCollection(SheetToRestoreFrom.Names, "xxx_BackedUpRangeAddress") Or _
              Not IsInCollection(SheetToRestoreFrom.Names, "xxx_VisibleRangeAddress") Or _
              Not IsInCollection(SheetToRestoreFrom.Names, "xxx_ValuesOnly") Then
5             CleanUpUndoBuffer SheetToRestoreFrom
6             GoTo EarlyExit
7         End If

8         BackedUpBookName = Evaluate(SheetToRestoreFrom.Names("xxx_BackedUpBookName").RefersTo)
9         BackedUpSheetName = Evaluate(SheetToRestoreFrom.Names("xxx_BackedUpSheetName").RefersTo)
10        BackedUpRangeAddress = Evaluate(SheetToRestoreFrom.Names("xxx_BackedUpRangeAddress").RefersTo)
11        VisibleRangeAddress = Evaluate(SheetToRestoreFrom.Names("xxx_VisibleRangeAddress").RefersTo)
12        ValuesOnly = Evaluate(SheetToRestoreFrom.Names("xxx_ValuesOnly").RefersTo)
13        If IsInCollection(SheetToRestoreFrom.Names, "xxx_RangeToSelectAfterUndoAddress") Then
14            RangeToSelectAfterUndoAddress = Evaluate(SheetToRestoreFrom.Names("xxx_RangeToSelectAfterUndoAddress").RefersTo)
15        End If

16        If Not IsInCollection(Application.Workbooks, BackedUpBookName) Then GoTo EarlyExit
17        Set wbTarget = Application.Workbooks(BackedUpBookName)

18        If Not IsInCollection(wbTarget.Worksheets, BackedUpSheetName) Then GoTo EarlyExit
19        Set wsTarget = wbTarget.Worksheets(BackedUpSheetName)

20        AreaAddresses = sTokeniseString(BackedUpRangeAddress)

21        Set SPH = CreateSheetProtectionHandler(wsTarget)
22        Set XSH = CreateExcelStateHandler(, , False)

23        If TypeName(Selection) = "Range" Then
24            Set OrigSelection = Selection
25        End If
26        If Not ActiveWindow Is Nothing Then
27            Set origViewPort = ActiveWindow.VisibleRange
28        End If

29        If ValuesOnly Then
30            For Each address In AreaAddresses
31                Set SourceRange = SheetToRestoreFrom.Range(address)
32                Set TargetRange = wsTarget.Range(address)
33                MyPaste TargetRange, SourceRange.Value2
34            Next address
35        Else
36            For Each address In AreaAddresses
37                Set SourceRange = SheetToRestoreFrom.Range(address)
38                Set TargetRange = wsTarget.Range(address)
39                SourceRange.Copy
40                Application.DisplayAlerts = False
                  'Have DisplayAlerts = False to avoid warning message: _
                  A formula or sheet you want to move or copy contains the name 'Foo' _
                                                                                 which already exists on the destination workbook etc. etc.
41                TargetRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
42            Next address
43        End If

44        If Not OrigSelection Is Nothing And Not origViewPort Is Nothing Then
45            Application.GoTo origViewPort
46            Application.GoTo OrigSelection
47        End If

48        Application.GoTo wsTarget.Range(VisibleRangeAddress)
49        wsTarget.Range(RangeToSelectAfterUndoAddress).Select
50        Application.CutCopyMode = False

EarlyExit:
51        CleanUpUndoBuffer SheetToRestoreFrom
52        Exit Sub
ErrHandler:
53        Throw "#RestoreRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

