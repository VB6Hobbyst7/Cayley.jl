Attribute VB_Name = "mod1DArrays"
Option Explicit
Option Private Module

'Almost all SolumAddin are designed to handle 2-dimensional arrays, but returns from JuliaExcel are likely 1-dimensional
'So implement some simple utilities in this module.

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Repeat
' Author     : Philip Swannell
' Date       : 17-Jan-2022
' Purpose    : Return is a one-dimensional array, indexed form 1 to n, each element is equal to x.
' -----------------------------------------------------------------------------------------------------------------------
Function Repeat(x, N As Long)
          Dim i As Long
          Dim Res() As Variant
1         On Error GoTo ErrHandler
2         ReDim Res(1 To N)
3         For i = 1 To N
4             Res(i) = x
5         Next
6         Repeat = Res
7         Exit Function
ErrHandler:
8         Throw "#Repeat (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Concatenate1DArrays
' Author     : Philip Swannell
' Date       : 18-Jan-2022
' Purpose    : Concatenate 1 dimensional arrays to a single 1-dimensional array. Lower bound of return is 1, irrespective
'              of lower bounds of input arrays.
' -----------------------------------------------------------------------------------------------------------------------
Function Concatenate1DArrays(ParamArray ArraysToConcat())

          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim N As Long
          Dim Res()

1         On Error GoTo ErrHandler
2         For i = LBound(ArraysToConcat) To UBound(ArraysToConcat)
3             N = N + UBound(ArraysToConcat(i)) - LBound(ArraysToConcat(i)) + 1
4         Next i

5         ReDim Res(1 To N)
6         For i = LBound(ArraysToConcat) To UBound(ArraysToConcat)
7             For j = LBound(ArraysToConcat(i)) To UBound(ArraysToConcat(i))
8                 k = k + 1
9                 Res(k) = ArraysToConcat(i)(j)
10            Next j
11        Next i

12        Concatenate1DArrays = Res

13        Exit Function
ErrHandler:
14        Throw "#Concatenate1DArrays (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDToTwoD
' Author     : Philip Swannell
' Date       : 18-Jan-2022
' Purpose    : Convert a one-dimensional array to 1-based, 1-column, 2-dimensional array.
' -----------------------------------------------------------------------------------------------------------------------
Function OneDToTwoD(x)
          Dim i As Long
          Dim offset As Long
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         ReDim Res(1 To UBound(x) - LBound(x) + 1, 1 To 1)
3         offset = 1 - LBound(x)

4         For i = LBound(x) To UBound(x)
5             Res(i + offset, 1) = x(i)
6         Next i

7         OneDToTwoD = Res

8         Exit Function
ErrHandler:
9         Throw "#OneDToTwoD (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
