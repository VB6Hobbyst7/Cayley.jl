VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHStacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsHStacker
' Author    : Philip Swannell
' Date      : 24 Aug 2020, amended version of clsStacker to do "side-by-side" stacking
' Purpose   : Class to implement "accumulation by stacking" of arrays of data. Class has
'             methods: StackData and Report. Restriction: Every call to Stack must pass in an
'             array with the same number of rows as previous calls - that number is initialised
'             in the first call to Stack.
'Data passed to StackData can be singleton (Long, Double, Boolean, String, Date etc) or 1-d array or 2-d array either 1-based or 0-based.
'             Return from Report is always a two dimensional array (1-based)
' 16-May-2019 Added methods Stack0D, Stack1D and Stack2D. Faster for stacking data when we know dimensionality.

' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_TheArray() As Variant
Private m_NumCols As Long
Private m_NumRows As Long
Private m_LastWrittenCol As Long

Public Function Report()
1         On Error GoTo ErrHandler

2         If m_LastWrittenCol = 0 Then
3             Report = "#Nothing to report!"
4         Else
5             If m_LastWrittenCol = m_NumCols Then
6                 Report = m_TheArray
7             Else
8                 Report = sSubArray(m_TheArray, , , , m_LastWrittenCol)
9             End If
10        End If

11        Exit Function
ErrHandler:
12        Throw "#clsHStacker.Report (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Public Function NumCols()
1         NumCols = m_LastWrittenCol
End Function

Private Function Max(x As Long, y As Long)
1         If x > y Then
2             Max = x
3         Else
4             Max = y
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StackData
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Use to Stack data of either 0,1 or 2 dimensions
' -----------------------------------------------------------------------------------------------------------------------
Public Sub StackData(DataToStack As Variant)
1         On Error GoTo ErrHandler
2         Select Case NumDimensions(DataToStack)
              Case 0
3                 Stack0D DataToStack
4             Case 1
5                 Stack1D DataToStack
6             Case 2
7                 Stack2D DataToStack
8             Case Else
9                 Throw "Too many dimensions to stack!"
10        End Select
11        Exit Sub
ErrHandler:
12        Throw "#clsHStacker.StackData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack0D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking a singleton
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack0D(DataToStack As Variant)

1         On Error GoTo ErrHandler

2         If m_NumRows = 0 Then m_NumRows = 1

3         If 1 <> m_NumRows Then Throw "Inconsistent number of rows in DataToStack"
4         If m_LastWrittenCol + 1 > m_NumCols Then
5             ReDim Preserve m_TheArray(1 To m_NumRows, 1 To Max(2 * m_NumCols, m_NumCols + 1))
6             m_NumCols = UBound(m_TheArray, 2)
7         End If

8         m_TheArray(1, m_LastWrittenCol + 1) = DataToStack
9         m_LastWrittenCol = m_LastWrittenCol + 1

10        Exit Sub
ErrHandler:
11        Throw "#clsHStacker.Stack0D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack1D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking 1-dimensional array
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack1D(DataToStack As Variant)
          Dim Base As Long
          Dim j As Long
          Dim NumRows As Long

1         On Error GoTo ErrHandler

2         Base = LBound(DataToStack)
3         NumRows = UBound(DataToStack) - Base + 1

4         If m_NumRows = 0 Then m_NumRows = NumRows

5         If NumRows <> m_NumRows Then Throw "Inconsistent number of rows in DataToStack"
6         If m_LastWrittenCol + 1 > m_NumCols Then
7             ReDim Preserve m_TheArray(1 To m_NumRows, 1 To Max(2 * m_NumCols, m_NumCols + 1))
8             m_NumCols = UBound(m_TheArray, 2)
9         End If

10        For j = 1 To NumRows
11            m_TheArray(j, m_LastWrittenCol + 1) = DataToStack(j + Base - 1)
12        Next j
13        m_LastWrittenCol = m_LastWrittenCol + 1

14        Exit Sub
ErrHandler:
15        Throw "#clsHStacker.Stack1D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack2D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking 2-dimensional array
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack2D(DataToStack As Variant)
          Dim Base As Long
          Dim i As Long
          Dim j As Long
          Dim NumCols As Long
          Dim NumRows

1         On Error GoTo ErrHandler

2         Base = LBound(DataToStack, 1)
3         NumRows = UBound(DataToStack, 1) - Base + 1
4         NumCols = UBound(DataToStack, 2) - Base + 1

5         If m_NumRows = 0 Then m_NumRows = NumRows

6         If NumRows <> m_NumRows Then Throw "Inconsistent number of rows in DataToStack"
7         If m_LastWrittenCol + NumCols > m_NumCols Then
8             ReDim Preserve m_TheArray(1 To m_NumRows, 1 To Max(2 * m_NumCols, m_NumCols + NumCols))
9             m_NumCols = UBound(m_TheArray, 2)
10        End If

11        For i = 1 To NumRows
12            For j = 1 To NumCols
13                m_TheArray(i, m_LastWrittenCol + j) = DataToStack(i + Base - 1, j + Base - 1)
14            Next j
15        Next i
16        m_LastWrittenCol = m_LastWrittenCol + NumCols

17        Exit Sub
ErrHandler:
18        Throw "#clsHStacker.Stack2D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

