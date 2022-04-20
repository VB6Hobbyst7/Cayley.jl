Attribute VB_Name = "modEnum"
Option Explicit

Enum EnmOptStyle
    OptStyleCall = 1
    OptStylePut = 2
    OptStyleBuy = 3
    OptStyleSell = 4
    optStyleUpDigital = 5
    optStyleDownDigital = 6
End Enum
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StringToOptStyle
' Author    : Philip Swannell
' Date      : 07-Jul-2015
' Purpose   : Function to convert user-friendly strings into the enumeration EnmOptStyle
'             as recognised by sBlackScholes, bscore, sFxOptionPFE etc.
'             This should be the only place we hard-code the strings "Put", "Call", "UpDigital" etc.
' -----------------------------------------------------------------------------------------------------------------------
Function StringToOptStyle(ByVal OptionStyle As String, AllowForward As Boolean) As EnmOptStyle

          Const ErrString1 = "OptionStyle must be Call (or C), Put (or P), Buy (or B), Sell (or S), Up Digital (or UD), Down Digital (or DD)"
          Const ErrString2 = "OptionStyle must be Call (or C), Put (or P), Forward (or F), Up Digital (or UD), Down Digital (or DD)"

1         OptionStyle = UCase$(Replace(CStr(OptionStyle), " ", vbNullString))

2         Select Case OptionStyle
              Case "B", "BUY"
3                 If AllowForward Then Throw ErrString2
4                 StringToOptStyle = OptStyleBuy
5             Case "S", "SELL", CStr(OptStyleSell)
6                 If AllowForward Then Throw ErrString2
7                 StringToOptStyle = OptStyleSell
8             Case "F", "FORWARD"
9                 If Not AllowForward Then Throw ErrString1
10                StringToOptStyle = OptStyleBuy
11            Case "C", "CALL"
12                StringToOptStyle = OptStyleCall
13            Case "P", "PUT"
14                StringToOptStyle = OptStylePut
15            Case "UD", "UPDIGITAL"
16                StringToOptStyle = optStyleUpDigital
17            Case "DD", "DOWNDIGITAL"
18                StringToOptStyle = optStyleDownDigital
19            Case Else
20                Throw IIf(AllowForward, ErrString2, ErrString1)
21        End Select
22        Exit Function
End Function

Function StringsToOptStyle(OptionStyles As Variant, AllowForward As Boolean)
          Dim i As Long
          Dim j As Long
          Dim ND As Long
          Dim Result() As EnmOptStyle
1         On Error GoTo ErrHandler
2         If VarType(OptionStyles) < vbArray Then
3             StringsToOptStyle = StringToOptStyle(CStr(OptionStyles), AllowForward)
4             Exit Function
5         End If
6         ND = NumDimensions(OptionStyles)
7         If ND = 1 Then
8             ReDim Result(LBound(OptionStyles) To UBound(OptionStyles))
9             For i = LBound(OptionStyles) To UBound(OptionStyles)
10                Result(i) = StringToOptStyle(OptionStyles(i), AllowForward)
11            Next i
12            StringsToOptStyle = Result
13        ElseIf ND = 2 Then
14            ReDim Result(LBound(OptionStyles, 1) To UBound(OptionStyles, 1), LBound(OptionStyles, 2) To UBound(OptionStyles, 2))
15            For i = LBound(OptionStyles, 1) To UBound(OptionStyles, 1)
16                For j = LBound(OptionStyles, 2) To UBound(OptionStyles, 2)
17                    Result(i, j) = StringToOptStyle(CStr(OptionStyles(i, j)), AllowForward)
18                Next j
19            Next i
20            StringsToOptStyle = Result
21        End If
22        Exit Function
ErrHandler:
23        Throw "#StringsToOptStyle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
