Attribute VB_Name = "modArrayString"
Option Explicit
Private Const DQ = """"
Private Const TwoDQ = """"""
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMakeArrayString
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : Converts an array of arbitrary data into a single string, which can be useful for data
'             transfer. The inverse of this function is sParseArrayString.
' Arguments
' TheArray  : An arbitrary array of values.
' -----------------------------------------------------------------------------------------------------------------------
Function sMakeArrayString(TheArray As Variant)
Attribute sMakeArrayString.VB_Description = "Converts an array of arbitrary data into a single string, which can be useful for data transfer. The inverse of this function is sParseArrayString."
Attribute sMakeArrayString.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim TempArray() As String

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, NR, NC
3         ReDim TempArray(1 To 2 * NR * NC + 1)

4         k = 1
5         TempArray(1) = "{"
6         For i = 1 To NR
7             For j = 1 To NC
8                 k = k + 1
9                 TempArray(k) = SingletonToText(TheArray(i, j))
10                k = k + 1
11                TempArray(k) = IIf(j = NC, ";", ",")
12            Next j
13        Next i
14        TempArray(2 * NR * NC + 1) = "}"
15        sMakeArrayString = VBA.Join(TempArray, vbNullString)
16        Exit Function
ErrHandler:
17        sMakeArrayString = "#sMakeArrayString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sParseArrayString
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : This function recreates an array from an array string made by sMakeArrayString.
' Arguments
' ArrayString: A string as can be returned by sMakeArrayString.
' -----------------------------------------------------------------------------------------------------------------------
Function sParseArrayString(ArrayString As String)
Attribute sParseArrayString.VB_Description = "This function recreates an array from an array string made by sMakeArrayString."
Attribute sParseArrayString.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim CommaCount As Long
          Dim DelimiterPositions() As Long
          Dim DQCount As Long
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As Variant
          Dim SemiColonCount As Long
          Const MalFormedError = "ArrayString must be valid output from sMakeArrayString"
          Dim FirstSemiAt As Long

1         On Error GoTo ErrHandler

2         If Left$(ArrayString, 1) <> "{" Then Throw MalFormedError
3         If Right$(ArrayString, 1) <> "}" Then Throw MalFormedError

4         For i = 1 To Len(ArrayString)
5             Select Case (Mid$(ArrayString, i, 1))
                  Case DQ
6                     DQCount = DQCount + 1
7                 Case ","
8                     If DQCount Mod 2 = 0 Then CommaCount = CommaCount + 1
9                 Case ";"
10                    If DQCount Mod 2 = 0 Then
11                        SemiColonCount = SemiColonCount + 1
12                        If FirstSemiAt = 0 Then
13                            FirstSemiAt = CommaCount + 1
14                        Else
15                            If (CommaCount + SemiColonCount) Mod FirstSemiAt <> 0 Then Throw MalFormedError
16                        End If
17                    End If
18            End Select
19        Next i
20        If DQCount Mod 2 <> 0 Then Throw MalFormedError

21        NR = SemiColonCount + 1
22        NC = (CommaCount + SemiColonCount + 1) / NR
23        If Not (NR * NC = CommaCount + SemiColonCount + 1) Then Throw MalFormedError

24        ReDim DelimiterPositions(1 To NR * NC + 1)

25        DelimiterPositions(1) = 1
26        DQCount = 0: k = 1
27        For i = 2 To Len(ArrayString) - 1
28            Select Case (Mid$(ArrayString, i, 1))
                  Case DQ
29                    DQCount = DQCount + 1
30                Case ",", ";"
31                    If DQCount Mod 2 = 0 Then
32                        k = k + 1
33                        DelimiterPositions(k) = i
34                    End If
35            End Select
36        Next i
37        DelimiterPositions(NR * NC + 1) = Len(ArrayString)

38        ReDim Result(1 To NR, 1 To NC)
39        k = 0
40        For i = 1 To NR
41            For j = 1 To NC
42                k = k + 1
43                Result(i, j) = TextToSingleton(Mid$(ArrayString, DelimiterPositions(k) + 1, DelimiterPositions(k + 1) - DelimiterPositions(k) - 1))
44            Next j
45        Next i

46        sParseArrayString = Result

47        Exit Function
ErrHandler:
48        sParseArrayString = "#sParseArrayString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SingletonToText
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : Sub-routine of sMakeArrayString and obverse of TextToSingleton
' -----------------------------------------------------------------------------------------------------------------------
Private Function SingletonToText(x As Variant) As String
          Dim Res As String
1         On Error GoTo ErrHandler
2         If IsError(x) Then
3             Select Case CStr(x)
                  Case "Error 2007"
4                     Res = "#DIV/0!"
5                 Case "Error 2029"
6                     Res = "#NAME?"
7                 Case "Error 2023"
8                     Res = "#REF!"
9                 Case "Error 2036"
10                    Res = "#NUM!"
11                Case "Error 2000"
12                    Res = "#NULL!"
13                Case "Error 2042"
14                    Res = "#N/A"
15                Case "Error 2015"
16                    Res = "#VALUE!"
17                Case "Error 2045"
18                    Res = "#SPILL!"
19                Case "Error 2047"
20                    Res = "#BLOCKED!"
21                Case "Error 2046"
22                    Res = "#CONNECT!"
23                Case "Error 2048"
24                    Res = "#UNKNOWN!"
25                Case "Error 2043"
26                    Res = "#GETTING_DATA!"
27                Case "Error 2049"
28                    Res = "#FIELD!"
29                Case "Error 2050"
30                    Res = "#CALC!"
31                Case Else
32                    Res = CStr(x)        'should never hit this line...
33            End Select
34        ElseIf VarType(x) = vbString Then
35            Res = DQ + Replace(x, DQ, TwoDQ) + DQ
36        ElseIf VarType(x) = vbBoolean Then
37            Res = UCase$(CStr(x))
38        Else
39            Res = SafeCStr(x)
40        End If
41        SingletonToText = Res
42        Exit Function
ErrHandler:
43        Throw "#SingletonToText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TextToSingleton
' Author    : Philip Swannell
' Date      : 25-Sep-2015
' Purpose   : Sub-routine of sParseArrayString and obverse of SingletonToText
' -----------------------------------------------------------------------------------------------------------------------
Private Function TextToSingleton(x As String)
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         If Left$(x, 1) = "#" Then
3             Select Case x
                  Case "#DIV/0!"
4                     Res = CVErr(xlErrDiv0)
5                 Case "#NAME?"
6                     Res = CVErr(xlErrName)
7                 Case "#REF!"
8                     Res = CVErr(xlErrRef)
9                 Case "#NUM!"
10                    Res = CVErr(xlErrNum)
11                Case "#NULL!"
12                    Res = CVErr(xlErrNull)
13                Case "#N/A"
14                    Res = CVErr(xlErrNA)
15                Case "#VALUE!"
16                    Res = CVErr(xlErrValue)
17                Case "#SPILL!"
18                    Res = CVErr(2045)    'CVErr(xlErrNoSpill)'These constants introduced in Excel 2016
19                Case "#BLOCKED!"
20                    Res = CVErr(2047)    'CVErr(xlErrBlocked)
21                Case "#CONNECT!"
22                    Res = CVErr(2046)    'CVErr(xlErrConnect)
23                Case "#UNKNOWN!"
24                    Res = CVErr(2048)    'CVErr(xlErrUnknown)
25                Case "#GETTING_DATA!"
26                    Res = CVErr(2043)    'CVErr(xlErrGettingData)
27                Case "#FIELD!"
28                    Res = CVErr(2049)    'CVErr(xlErrField)
29                Case "#CALC!"
30                    Res = CVErr(2050)    'CVErr(xlErrCalc)
31                Case Else
32                    Res = x
33            End Select
34        ElseIf Left$(x, 1) = DQ Then
35            Res = Replace(Mid$(x, 2, Len(x) - 2), TwoDQ, DQ)
36        ElseIf UCase$(x) = "TRUE" Then
37            Res = True
38        ElseIf UCase$(x) = "FALSE" Then
39            Res = False
40        ElseIf x = vbNullString Then
41            Res = Empty
42        Else
43            Res = SafeCDbl(x)
44        End If
45        TextToSingleton = Res
46        Exit Function
ErrHandler:
47        Throw "#TextToSingleton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SafeCStr(x As Variant)
1         On Error GoTo ErrHandler

2         SafeCStr = CStr(x)
3         Exit Function
ErrHandler:
4         SafeCStr = "Cannot represent variable of type " + TypeName(x) + " as a string"
End Function

Private Function SafeCDbl(x As String)
1         On Error GoTo ErrHandler
2         SafeCDbl = CDbl(x)
3         Exit Function
ErrHandler:
4         Throw "#SafeCDbl (line " & CStr(Erl) + "): " & " Cannot interpret string " + x + "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayMakeText
' Author    : Philip Swannell
' Date      : 28-Sep-2015
' Purpose   : Elementwise conversion of an array to strings. Strings are unchanged, 123 becomes the
'             string "123", logicals become either "TRUE" or "FALSE", and missings and
'             empties become empty strings. Errors are converted to strings such as
'             "#NAME?" or "#REF!".
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayMakeText(ByVal TheArray)
Attribute sArrayMakeText.VB_Description = "Element-wise conversion of an array to strings. Strings are unchanged, 123 becomes the string ""123"", logicals become either ""TRUE"" or ""FALSE"", and missings and empties become empty strings. Errors are converted to strings such as ""#NAME?"" or ""#REF!""."
Attribute sArrayMakeText.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As String

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, NR, NC

3         ReDim Result(1 To NR, 1 To NC)
4         For i = 1 To NR
5             For j = 1 To NC
6                 Select Case VarType(TheArray(i, j))
                      Case vbString
7                         Result(i, j) = TheArray(i, j)
8                     Case Else
9                         Result(i, j) = SingletonToText(TheArray(i, j))
10                End Select
11            Next j
12        Next i
13        sArrayMakeText = Result
14        Exit Function
ErrHandler:
15        sArrayMakeText = "#sArrayMakeText (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayExcelString
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Prepends all strings in TheArray with a single quote character. Non-string entries are
'             returned unchanged. The function is suitable for use in VBA to be on the
'             right-hand side of a Range.Value assignment, to avoid conversion of strings
'             to non-strings. But see also method MyPaste.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayExcelString(ByVal TheArray)
Attribute sArrayExcelString.VB_Description = "Prepends all strings in TheArray with a single quote character, leaving non-string unchanged. Use in VBA on the right-hand side of a Range.Value assignment, to avoid conversion of strings to non-strings."
Attribute sArrayExcelString.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, NR, NC
3         For i = 1 To NR
4             For j = 1 To NC
5                 If VarType(TheArray(i, j)) = vbString Then
6                     TheArray(i, j) = "'" + TheArray(i, j)
7                 End If
8             Next j
9         Next i
10        sArrayExcelString = TheArray
11        Exit Function
ErrHandler:
12        sArrayExcelString = "#sArrayExcelString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
