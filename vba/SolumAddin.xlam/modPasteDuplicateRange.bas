Attribute VB_Name = "modPasteDuplicateRange"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modPasteDuplicateRange
' Author    : Philip Swannell
' Date      : 08-Nov-2013
' Purpose   : Implements an operation that is a hybrid of copy-paste and cut-paste, and
'             we call it "Paste Duplicate Range". The paste area ends up looking as though
'             the copy area had been moved there but the copy area is left unchanged.
'             Formulas in the paste area are identical to those in the copy area except where
'             formulas reference cells inside the copy area. These formulas are updated to reference
'             a cell in the same relative address.
'         Implementation is simple:
'             1) Copy area to a temporary sheet (maintaining its cell address)
'             2) Move the copied range inside the temporary sheet to the correct cell address
'             3) Copy back to the active sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteDuplicateRange()
          Dim CopyOfErr As String
          Dim ESH As clsExcelStateHandler
          Dim ExpandedRange As Range
          Dim HaveMadeChanges As Boolean
          Dim Prompt As String
          Dim SourceInTemp As Range
          Dim SourceRange As Range
          Dim SUH As clsScreenUpdateHandler
          Dim TargetInTemp As Range
          Dim TargetRange As Range
          Dim tempBook As Excel.Workbook
          Dim TempSheet As Worksheet
          Const Title = "Paste Duplicate Range"
          Dim S As Shape

1         On Error GoTo ErrHandler
2         If ActiveWindow Is Nothing Then Throw "Please select a cell to copy to."
3         Set ESH = CreateExcelStateHandler(SetEnableEventsTo:=False)
4         Set SUH = CreateScreenUpdateHandler()

5         On Error Resume Next
6         GetCopiedRange SourceRange
7         On Error GoTo ErrHandler
8         If SourceRange Is Nothing Then Throw "You must first copy a range (Ctrl C or Shift Insert) before you can paste it as a duplicate range.", True

9         If SourceRange.Areas.Count > 1 Then
10            Application.GoTo SourceRange
11            Throw "That command cannot be used on multiple selections.", True
12        End If

13        Set ExpandedRange = ExpandRangeToIncludeEntireArrayFormulas(SourceRange)

14        If ExpandRangeToIncludeEntireArrayFormulas(SourceRange).address <> SourceRange.address Then
15            Prompt = "You copied range " + AddressND(SourceRange) + " but that has been expanded to range " + AddressND(ExpandedRange) + " to include entire array formulas." + vbLf + vbLf + "Do you want to continue?"
16            ExpandedRange.Copy
17            Set SourceRange = ExpandedRange
18            Set S = HighLightRangeWithShape(ExpandedRange)
19            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, "Yes - continue", "No - do nothing") <> vbOK Then
20                S.Delete
21                Exit Sub
22            Else
23                S.Delete
24            End If
25        End If

26        Set TargetRange = ActiveWindow.RangeSelection
27        If TargetRange.Areas.Count > 1 Then Throw "That command cannot be used on multiple selections.", True
28        If Not TargetRange.Parent Is SourceRange.Parent Then Throw "The cells cannot be pasted to a different sheet", True
29        If TargetRange.Cells.CountLarge > 1 And (TargetRange.Rows.Count <> SourceRange.Rows.Count Or _
              TargetRange.Columns.Count <> SourceRange.Columns.Count) Then
30            Throw "The cells cannot be pasted because the copy area and paste area " & _
                  "are not the same size and shape. Try one of the following:" + vbLf + vbLf + _
                  Chr$(149) & " Click a single cell, and then use Paste Duplicate Range." + vbLf + _
                  Chr$(149) & " Select a rectangle that's the same size and shape, and then use Paste Duplicate Range.", True
31        End If
32        If TargetRange.row + SourceRange.Rows.Count - 1 > TargetRange.Parent.Rows.Count Or _
              TargetRange.Column + SourceRange.Columns.Count - 1 > TargetRange.Parent.Columns.Count Then
33            Throw "Cannot copy to cells beyond the edges of the worksheet.", True
34        End If
35        Set TargetRange = TargetRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
36        If TargetRange.Parent Is SourceRange.Parent Then
37            If Not Application.Intersect(TargetRange, SourceRange) Is Nothing Then
38                Throw "The cells cannot be pasted because the copy and paste areas overlap.", True
39            End If
40        End If
41        If Not UnprotectAsk(TargetRange.Parent, Title) Then Exit Sub

          '   TestRangeForPartialArrayFormulas SourceRange, "copy area"  'No longer necessary
42        TestRangeForPartialArrayFormulas TargetRange, "paste area"

43        If TargetRange.Cells.CountLarge > MAX_CELLS_FOR_UNDO Then
44            Prompt = "Undo (Ctrl+Z) is not available for an operation as large as this." + vbLf + vbLf + _
                  "Continue without Undo?"
45            If MsgBoxPlus(Prompt, vbYesNo + vbDefaultButton2 + vbQuestion, Title) <> vbYes Then
46                Exit Sub
47            End If
48            CleanUpUndoBuffer shUndo
49        Else
50            BackUpRange TargetRange, shUndo
51        End If

52        Set tempBook = Application.Workbooks.Add
53        Set TempSheet = tempBook.Worksheets(1)
54        Set SourceInTemp = TempSheet.Range(SourceRange.address)
55        Set TargetInTemp = TempSheet.Range(TargetRange.address)
56        SourceRange.Copy SourceInTemp
57        SourceInTemp.Cut TargetInTemp
58        Set TargetInTemp = TempSheet.Range(TargetRange.address)
59        TargetInTemp.Copy TargetRange
60        HaveMadeChanges = True

61        tempBook.Close False
62        SourceRange.Copy
63        TargetRange.Select

64        Application.OnUndo "Undo " & Title, ThisWorkbook.Name & "!RestoreRange"

65        Exit Sub
ErrHandler:
66        CopyOfErr = "#PasteDuplicateRange (line " & CStr(Erl) + "): " & Err.Description & "!"
67        If Not tempBook Is Nothing Then tempBook.Close False
68        If HaveMadeChanges Then RestoreRange
69        SomethingWentWrong CopyOfErr, vbExclamation, Title
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetCopiedRange
' Author    : Philip Swannell
' Date      : 08-Nov-2013
' Purpose   : Returns a Range object of the currently copied selection, or throws an
'             error if there isn't one. This method is called from the SCRiPT workbook,
'             hence no longer Private (PGS 16-3-16)
' 10-Jan-19   Converted to Sub as part of effort to reduce exposure of functions to Excel
' -----------------------------------------------------------------------------------------------------------------------
Public Sub GetCopiedRange(ByRef TheRange As Range)    'DO NOT MAKE Private!
          Dim Cell1 As Range
          Dim Cell2 As Range
          Dim ConvexHull As Range
          Dim CopyOfErr As String
          Dim CountRepeatsRet As Variant
          Dim ExclamIsAt As Long
          Dim Format As Variant
          Dim Formats As Variant
          Dim Formula1 As String
          Dim Formula2 As String
          Dim i As Long
          Dim R1C1FormulasInLeftCol As Variant
          Dim R1C1FormulasInTopRow As Variant
          Dim SUH As clsScreenUpdateHandler
          Dim tempBook As Excel.Workbook
          Dim TempRange As Range
          Dim ThisArea As Range

1         On Error GoTo ErrHandler

2         If Application.CutCopyMode <> xlCopy Then
3             Throw "#No copied Range found!"
4             Exit Sub
5         End If
          'Examine ClipBoard formats to check that what's copied is indeed a range
          'Found this tip at http://www.ozgrid.com/forum/showthread.php?t=66773
6         Formats = Application.ClipboardFormats
7         For Each Format In Formats
8             If Format = xlClipboardFormatCSV Then
9                 GoTo Continue
10            End If
11        Next
12        Throw "#No copied Range found!"
13        Exit Sub
Continue:
14        Set SUH = CreateScreenUpdateHandler()
15        Set tempBook = Application.Workbooks.Add
16        tempBook.Worksheets(1).Paste Link:=True
17        Set TempRange = Selection

18        With TempRange
19            Formula1 = .Cells(1, 1).Formula
20            Formula2 = .Cells(.Rows.Count, .Columns.Count).Formula
21        End With

          'Rubberduck (2.4.1.4627) incorrectly flags these three lines as implicitly referencing the active sheet
22        Set Cell1 = Range(Right$(Formula1, Len(Formula1) - 1))
23        Set Cell2 = Range(Right$(Formula2, Len(Formula2) - 1))
24        Set ConvexHull = Range(Cell1, Cell2)
          'https://en.wikipedia.org/wiki/Convex_hull

25        If ConvexHull.Cells.CountLarge = TempRange.Cells.CountLarge Then
              ' Copied Range had one area only.
26            Set TheRange = ConvexHull
27        Else
              'There are now two possibilities: _
              a) Copied range had multiple areas, each of the same width and all aligned vertically; or _
                  b) Copied range had multiple areas, each of the same height and all aligned horizontally. _
                  It's not possible to copy other layouts of multiple-area ranges (as of Office 2013)

28            If ConvexHull.Rows.Count > TempRange.Rows.Count Then
                  'We're in case a)
29                ExclamIsAt = InStrRev(TempRange.Cells(1, 1).FormulaR1C1, "!")
30                R1C1FormulasInLeftCol = sReshape(vbNullString, TempRange.Rows.Count, 1)
31                For i = 1 To TempRange.Rows.Count
32                    R1C1FormulasInLeftCol(i, 1) = Mid$(TempRange.Cells(i, 1).FormulaR1C1, ExclamIsAt + 1)
33                Next i
34                CountRepeatsRet = sCountRepeats(R1C1FormulasInLeftCol, "CFH")
35                Set TheRange = Cell1        ' to initialise
36                For i = 1 To sNRows(CountRepeatsRet)
37                    If InStr(CountRepeatsRet(i, 1), "R[") > 0 Then
38                        Set ThisArea = ConvexHull.Rows(CountRepeatsRet(i, 2) + CLng(CoreStringBetweenStrings(CountRepeatsRet(i, 1), "R[", "]")) - ConvexHull.row + 1)
39                        Set ThisArea = ThisArea.Resize(CountRepeatsRet(i, 3))
40                    Else
                          'This is a strange case. When one of the areas of the copied range is a single cell the formula pasted at the paste-link step is in absolute rather than relative form
41                        Set ThisArea = ConvexHull.Rows(CLng(sStringBetweenStrings(CountRepeatsRet(i, 1), "R", "C")) - ConvexHull.row + 1)
42                    End If
43                    Set TheRange = Application.Union(TheRange, ThisArea)
44                Next i
45            Else
                  'We're in case b)
46                ExclamIsAt = InStrRev(TempRange.Cells(1, 1).FormulaR1C1, "!")
47                R1C1FormulasInTopRow = sReshape(vbNullString, TempRange.Columns.Count, 1)
48                For i = 1 To TempRange.Columns.Count
49                    R1C1FormulasInTopRow(i, 1) = Mid$(TempRange.Cells(1, i).FormulaR1C1, ExclamIsAt + 1)
50                Next i
51                CountRepeatsRet = sCountRepeats(R1C1FormulasInTopRow, "CFH")
52                Set TheRange = Cell1        ' to initialise
53                For i = 1 To sNRows(CountRepeatsRet)
54                    If InStr(CountRepeatsRet(i, 1), "C[") > 0 Then
55                        Set ThisArea = ConvexHull.Columns(CountRepeatsRet(i, 2) + CLng(CoreStringBetweenStrings(CountRepeatsRet(i, 1), "C[", "]")) - ConvexHull.Column + 1)
56                        Set ThisArea = ThisArea.Resize(, CountRepeatsRet(i, 3))
57                    Else
                          'The strange case described above
58                        Set ThisArea = ConvexHull.Columns(CLng(CoreStringBetweenStrings(CountRepeatsRet(i, 1), "C", vbNullString)) - ConvexHull.Column + 1)
59                    End If
60                    Set TheRange = Application.Union(TheRange, ThisArea)
61                Next i
62            End If
63        End If

64        tempBook.Close False

65        Exit Sub
ErrHandler:
66        CopyOfErr = "#GetCopiedRange (line " & CStr(Erl) + "): " & Err.Description & "!"
67        If Not tempBook Is Nothing Then tempBook.Close False
68        Throw CopyOfErr
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestRangeForPartialArrayFormulas
' Author    : Philip Swannell
' Date      : 13-May-2015
' Purpose   : argument NameForRangeToTest can take values "copy area", "paste area", "calculate"
'             and is used only to generate a helpful error message.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestRangeForPartialArrayFormulas(RangeToTest As Range, NameForRangeToTest)
          Dim BlankCellsInRangeToTest As Range
          Dim c As Range
          Dim FormulaCellsInRangeToTest As Range
          Dim nonEmptiesFound As Boolean

1         On Error GoTo ErrHandler
2         Set BlankCellsInRangeToTest = BlankCellsInRange(RangeToTest)
3         If BlankCellsInRangeToTest Is Nothing Then
4             nonEmptiesFound = True
5         Else
6             If Not RangesIdentical(BlankCellsInRangeToTest, RangeToTest) Then
7                 nonEmptiesFound = True
8             End If
9         End If
10        If nonEmptiesFound Then
11            Set FormulaCellsInRangeToTest = CellsWithFormulasInRange(RangeToTest)
12            If Not FormulaCellsInRangeToTest Is Nothing Then

13                For Each c In FormulaCellsInRangeToTest.Cells
14                    If c.HasArray Then
15                        If Application.Intersect(c.CurrentArray, RangeToTest).Cells.CountLarge <> c.CurrentArray.Cells.CountLarge Then
16                            On Error GoTo 0
17                            If NameForRangeToTest = "calculate" Then
18                                Throw "The selection cannot be calculated because the array formula at " & _
                                      AddressND(c.CurrentArray) & " intersects the selection. If the selection" & _
                                      " contains an array formula then make sure it contains all of the cells of the array formula.", True
19                            Else
20                                Throw "The cells cannot be pasted because the array formula at " & _
                                      AddressND(c.CurrentArray) & " intersects the " + _
                                      NameForRangeToTest + ". If the " + NameForRangeToTest + _
                                      " contains an array formula then make sure it contains all of the cells of the array formula.", True
21                            End If
22                        End If
23                    End If
24                Next c
25            End If
26        End If
27        Exit Sub
ErrHandler:
28        Throw "#TestRangeForIntersectingFormulas (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HighLightRangeWithShape
' Author     : Philip Swannell
' Date       : 16-Dec-2019
' Purpose    : Place a shape on top of a range, to draw the user's attention to the range. Returns the shape to allow for
'              subsequent deletion...
' -----------------------------------------------------------------------------------------------------------------------
Private Function HighLightRangeWithShape(ByVal R As Range) As Shape
          Dim S As Shape
          Dim EnlargedVisibleRange As Range
          Dim SPH As clsSheetProtectionHandler

1         Set SPH = CreateSheetProtectionHandler(R.Parent)

2         On Error GoTo ErrHandler

          Dim nudgeTop As Long, nudgeLeft As Long, nudgeHeight As Long, nudgeWidth As Long
          
3         If ActiveWindow Is Nothing Then
4             Set S = R.Parent.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
5             S.Visible = msoFalse
6             Set HighLightRangeWithShape = S
7             Exit Function
8         ElseIf Not R.Parent Is ActiveSheet Then
9             Set S = R.Parent.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
10            S.Visible = msoFalse
11            Set HighLightRangeWithShape = S
12            Exit Function
13        End If
          'There is some limit on the size of shapes, so we only put it above the intersection of R and a slightly enlarged VisibleRange, _
           this avoids getting "The specified value is out of range" error when setting the size of the shape.
14        With ActiveWindow.VisibleRange
15            nudgeTop = IIf(.row = 1, 0, -1)
16            nudgeLeft = IIf(.Column = 1, 0, -1)
17            nudgeHeight = IIf(.row + .Rows.Count - 1 = .Parent.Rows.Count, 0, 1) - nudgeTop
18            nudgeWidth = IIf(.Column + .Columns.Count - 1 = .Parent.Columns.Count, 0, 1) - nudgeLeft
19            Set EnlargedVisibleRange = .Offset(nudgeTop, nudgeLeft).Resize(.Rows.Count + nudgeHeight, .Columns.Count + nudgeWidth)
20        End With

21        If Application.Intersect(R, EnlargedVisibleRange) Is Nothing Then
22            Set S = R.Parent.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
23            S.Visible = msoFalse
24        Else
25            Set R = Application.Intersect(R, EnlargedVisibleRange)
26            Set S = R.Parent.Shapes.AddShape(msoShapeRectangle, R.Left, R.Top, R.Width, R.Height)
27            With S.Fill
28                .Visible = msoTrue
29                .ForeColor.ObjectThemeColor = msoThemeColorBackground2
30                .ForeColor.TintAndShade = 0
31                .ForeColor.Brightness = -0.5
32                .Transparency = 0.6899999976
33                .Solid
34            End With
35            With S.Line
36                .Visible = msoTrue
37                .ForeColor.ObjectThemeColor = msoThemeColorAccent6
38                .ForeColor.TintAndShade = 0
39                .ForeColor.Brightness = -0.25
40                .Transparency = 0
41                .Weight = 1.25
42                .DashStyle = msoLineSysDash
43            End With
44        End If

45        Set HighLightRangeWithShape = S

46        Exit Function
ErrHandler:
47        Throw "#HighLightRangeWithShape (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


