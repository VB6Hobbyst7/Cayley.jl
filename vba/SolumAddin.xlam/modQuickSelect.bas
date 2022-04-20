Attribute VB_Name = "modQuickSelect"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNthSmallest
' Author    : Philip Swannell
' Date      : 04-Jun-2015
' Purpose   : Returns the Nth smallest element of ArrayToSearch. If N is 1 the return is the smallest
'             element of ArrayToSearch, 2 for the second smallest etc. For details see
'             http://en.wikipedia.org/wiki/Quickselect.
'
'             The function duplicates Excel's function SMALL.
' Arguments
' ArrayToSearch: An array of numbers.
' N         : A whole number between 1 and the number of elements in ArrayToSearch.
' -----------------------------------------------------------------------------------------------------------------------
Function sNthSmallest(ByVal ArrayToSearch, N As Variant)
Attribute sNthSmallest.VB_Description = "Returns the Nth smallest element of ArrayToSearch. If N is 1 the return is the smallest element of ArrayToSearch, 2 for the second smallest etc. For details see http://en.wikipedia.org/wiki/Quickselect.\n\nThe function duplicates Excel's function SMALL."
Attribute sNthSmallest.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim c As Variant
          Dim ErrorString As String
          Dim NR As Long

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti ArrayToSearch, N
3         For Each c In ArrayToSearch
4             If Not IsNumberOrDate(c) Then Throw "ArrayToSearch must all be numbers"
5         Next c

6         If sNCols(ArrayToSearch) > 1 Then
7             If sNRows(ArrayToSearch) = 1 Then
8                 ArrayToSearch = sArrayTranspose(ArrayToSearch)
9             Else
10                ArrayToSearch = sReshape(ArrayToSearch, sNRows(ArrayToSearch) * sNCols(ArrayToSearch), 1)
11            End If
12        End If

13        NR = sNRows(ArrayToSearch)
14        If sNCols(ArrayToSearch) > 1 Then Throw "ArrayToSearch must have one column or one row"

15        ErrorString = "N must be a whole number between 1 and " + CStr(NR) + " or a column array of such numbers"
16        If sNCols(N) > 1 Then Throw ErrorString
17        For Each c In N
18            If Not IsNumber(c) Then Throw ErrorString
19            If Not CLng(c) = c Then Throw ErrorString
20            If c < 1 Or c > NR Then Throw ErrorString
21        Next c

22        sNthSmallest = QuickSelectMulti(ArrayToSearch, N)
23        Exit Function
ErrHandler:
24        sNthSmallest = "#sNthSmallest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSelectMulti
' Author    : Philip Swannell
' Date      : 06-Jun-2015
' Purpose   : A multi-call version of quick select. Will be more efficient if the Ns are monotonic
' -----------------------------------------------------------------------------------------------------------------------
Private Function QuickSelectMulti(ByRef List, NS As Variant)
          Dim i As Long
          Dim NumNs As Long
          Dim NumRows As Long
          Dim Res

1         On Error GoTo ErrHandler
2         NumNs = sNRows(NS)
3         NumRows = sNRows(List)

4         Res = sReshape(0, NumNs, 1)
5         For i = 1 To NumNs
6             If i = 1 Then
7                 Res(i, 1) = QuickSelect(List, 1, NumRows, CLng(NS(i, 1)))
8             ElseIf NS(i, 1) > NS(i - 1, 1) Then
9                 Res(i, 1) = QuickSelect(List, CLng(NS(i - 1, 1)), NumRows, CLng(NS(i, 1)))
10            ElseIf NS(i, 1) < NS(i - 1, 1) Then
11                Res(i, 1) = QuickSelect(List, 1, CLng(NS(i - 1, 1)), CLng(NS(i, 1)))
12            Else
13                Res(i, 1) = Res(i - 1, 1)
14            End If
15        Next i
16        QuickSelectMulti = Res
17        Exit Function
ErrHandler:
18        Throw "#QuickSelectMulti (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : QuickSelect
' Author    : Philip Swannell
' Date      : 04-Jun-2015
' Purpose   : Implementation of the well-known QuickSelect algorithm. See http://en.wikipedia.org/wiki/Quickselect.
'  Problem: Out-of-stackspace errors can occur if List has many repeated elements
' -----------------------------------------------------------------------------------------------------------------------
Private Function QuickSelect(List, Left As Long, Right As Long, N As Long)
          Dim PivotIndex As Long

1         On Error GoTo ErrHandler
2         If Left = Right Then
3             QuickSelect = List(Left, 1)
4             Exit Function
5         End If
          'We choose a new PivotIndex between Left and Right. Making random choice leads to "almost certain" linear execution time - see wikipedia article
6         PivotIndex = Left + Rnd() * (Right - Left + 1) - 0.5

7         PivotIndex = Partition(List, Left, Right, PivotIndex)
8         If N = PivotIndex Then
9             QuickSelect = List(N, 1)
10        ElseIf N < PivotIndex Then
11            QuickSelect = QuickSelect(List, Left, PivotIndex - 1, N)
12        Else
13            QuickSelect = QuickSelect(List, PivotIndex + 1, Right, N)
14        End If
15        Exit Function
ErrHandler:
16        Throw "#QuickSelect (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Partition
' Author    : Philip Swannell
' Date      : 04-Jun-2015
' Purpose   : Performs a partial sort on List. If PivotValue = List(PivotIndex) then
'             all elements less than PivotIndex appear before (above) PivotValue and all elements
'             greater appear after (below). Return from the function is the index of PivotValue
'             in the partially sorted list, which would also be its index in the fully sorted list.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Partition(ByRef List, Left, Right, PivotIndex)
          Dim i As Long
          Dim pivotValue As Double
          Dim StoreIndex

1         On Error GoTo ErrHandler
2         pivotValue = List(PivotIndex, 1)
3         Swap List, PivotIndex, Right
4         StoreIndex = Left
5         For i = Left To Right - 1
6             If List(i, 1) < pivotValue Then
7                 Swap List, StoreIndex, i
8                 StoreIndex = StoreIndex + 1
9             End If
10        Next i
11        Swap List, Right, StoreIndex
12        Partition = StoreIndex
13        Exit Function
ErrHandler:
14        Throw "#Partition (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Swap
' Author    : Philip Swannell
' Date      : 04-Jun-2015
' Purpose   : List is 2-d array with single column. Method swaps elements Indx1 and Indx2
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Swap(ByRef List, Indx1, Indx2)
          Dim tmp
1         On Error GoTo ErrHandler
2         tmp = List(Indx1, 1)
3         List(Indx1, 1) = List(Indx2, 1)
4         List(Indx2, 1) = tmp
5         Exit Sub
ErrHandler:
6         Throw "#Swap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
