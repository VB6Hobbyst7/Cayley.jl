Attribute VB_Name = "modPivot"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFilePivot
' Author    : Philip Swannell
' Date      : 12-Nov-2017
' Purpose   : Wraps function FilePivot in Pivot.R
' -----------------------------------------------------------------------------------------------------------------------
Function sFilePivot(FileName As String, FilterField1 As String, Filter1 As String, FilterField2 As String, Filter2 As String, _
        ColumnField As String, RowField As String, ValueFields, Optional ColumnOrder, Optional RowOrder, _
        Optional TotalsBeneath As Boolean = False, Optional TotalsToRight As Boolean = False)

          Dim Expression As String
          Static HaveCalledAlready As Boolean

1         On Error GoTo ErrHandler

2         If Not HaveCalledAlready Then
3             CheckR "sFilePivot", gPackagesSAI, gRSourcePath & "SolumAddin.R"
4             HaveCalledAlready = True
5         End If

          Const Sep As String = ""","""
6         Expression = "FilePivot2(""" + Replace(FileName, "\", "/") + Sep + FilterField1 + Sep + Filter1 + Sep + FilterField2 + Sep + Filter2 + Sep + _
              ColumnField + Sep + RowField + """," + ArrayToRLiteral(ValueFields) + "," + ArrayToRLiteral(ColumnOrder, "NULL") + "," + ArrayToRLiteral(RowOrder, "NULL") + _
              "," + UCase$(TotalsBeneath) + "," + UCase$(TotalsToRight) + ")"
7         Debug.Print Expression
8         sFilePivot = sExecuteRCode(Expression)

9         Exit Function
ErrHandler:
10        sFilePivot = "#sPivot (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
