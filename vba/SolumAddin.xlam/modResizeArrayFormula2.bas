Attribute VB_Name = "modResizeArrayFormula2"
Option Explicit
'PGS 19 Dec 2020 modResizeArrayFormula is a PrivateModule, but turns out that the Cayley workbook was using some of its methods, so moved them to this not Private Module.

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IWC - subroutine of IntersectWithComplement
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : Returns a Range object which is all cells in RangeA not in RangeB.
'             RangeB assumed to have one area only and be on the same sheet as RangeA
'             Is there an easier way to do this?
' -----------------------------------------------------------------------------------------------------------------------
Function IWC(RangeA As Range, RangeB As Range)
          Dim NC As Long
          Dim NR As Long
          Dim Result As Range
          Dim Result1 As Range
          Dim Result2 As Range
          Dim Result3 As Range
          Dim Result4 As Range

1         On Error GoTo ErrHandler

2         NR = RangeB.Parent.Rows.Count
3         NC = RangeB.Parent.Columns.Count

4         If RangeB.Column > 1 Then
5             Set Result1 = Application.Intersect(RangeA, RangeB.Parent.Cells(1, 1).Resize(1, RangeB.Column - 1).EntireColumn)
6         End If

7         If RangeB.Column + RangeB.Columns.Count - 1 < NC Then
8             Set Result2 = Application.Intersect(RangeA, RangeB.Cells(1, RangeB.Columns.Count + 1).Resize(1, NC - RangeB.Column - RangeB.Columns.Count + 1).EntireColumn)
9         End If

10        If RangeB.row > 1 Then
11            Set Result3 = Application.Intersect(RangeA, RangeB.Parent.Cells(1, 1).Resize(RangeB.row - 1, 1).EntireRow)
12        End If

13        If RangeB.row + RangeB.Rows.Count - 1 < NR Then
14            Set Result4 = Application.Intersect(RangeA, RangeB.Cells(RangeB.Rows.Count + 1, 1).Resize(NR - RangeB.row - RangeB.Rows.Count + 1, 1).EntireRow)
15        End If

16        Set Result = UnionOfRanges(Result1, Result2, Result3, Result4)

17        Set IWC = Result

18        Exit Function
ErrHandler:
19        Throw "#IWC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IntersectWithComplement. This function is called from the Cayley workbook.
' Author    : Philip Swannell
' Date      : 26-Jun-2013
' Purpose   : Returns a Range object which is all cells in RangeA not in RangeB.
'             RangeA and RangeB can be multi-area, but must be on the same sheet.
'             Returns Nothing when RangeA is entirely inside RangeB - see function sRangeContainsRange
' -----------------------------------------------------------------------------------------------------------------------
Function IntersectWithComplement(RangeA As Range, RangeB As Range) As Range
          Dim a As Range
          Dim Result As Range
          Dim ThisPart As Range

1         On Error GoTo ErrHandler

2         For Each a In RangeB.Areas
3             Set ThisPart = IWC(RangeA, a)
4             If ThisPart Is Nothing Then
5                 Exit Function
6             End If

7             If Not Result Is Nothing Then
8                 Set Result = Application.Intersect(Result, ThisPart)
9                 If Result Is Nothing Then
10                    Exit Function
11                End If
12            Else
13                Set Result = ThisPart
14            End If
15        Next a
16        If Not Result Is Nothing Then
17            Set IntersectWithComplement = Result
18        End If
19        Exit Function
ErrHandler:
20        Throw "#IntersectWithComplement (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
