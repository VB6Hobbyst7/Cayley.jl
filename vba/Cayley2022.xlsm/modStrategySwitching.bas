Attribute VB_Name = "modStrategySwitching"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modStrategySwitching
' Author    : Philip Swannell
' Date      : 01-Sep-2015
' Purpose   : Code to handle hadging strategies that vary as the amount of trade headroom changes
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetOptionsStrategy
' Author    : Philip Swannell
' Date      : 01-Sep-2015
' Purpose   : Motivation: Airbus want to see the effect of a hedging strategy that varies
'             as lines become tight. To implement this we allow the parameters of the
'             OptionsStrategy (ForwardsRation thru CallStrikeOffset) to be comma-separated
'             lists instead of numbers. A further "StrategySwitchPoints" parameter facilitates
'             switching between the elements of (say) ForwardsRatio according to how much
'             trade headroom remains, or according to time if first token of StrategySwitchPoints
'             is "SwitchOnTime"
' -----------------------------------------------------------------------------------------------------------------------
Function SetOptionsStrategy(ByVal TradeHeadroomInBillions As Double, TimeInMonths As Long, _
          ByVal ForwardsRatio As Variant, ByVal PutRatio As Variant, _
          ByVal CallRatio As Variant, ByVal PutStrikeOffset As Variant, CallStrikeOffset, _
          ByVal StrategySwitchPoints As Variant, ByRef dForwardsRatio As Double, ByRef dPutRatio As Double, _
          ByRef dCallRatio As Double, ByRef dPutStrikeOffset As Double, ByRef dCallStrikeOffset As Double)

1         On Error GoTo ErrHandler

          Const SwitchPointsErrorString = "StrategySwitchPoints must be a number or a comma separated list of " & _
              "numbers, or the text ""SwitchOnTime,"" followed by a comma separated list of numbers"

2         If IsEmpty(StrategySwitchPoints) Then
              'Simple case - there is no strategy switching
3             If Not IsNumber(ForwardsRatio) Then
4                 Throw "ForwardsRatio should be a number (when StrategySwitchPoints is empty)"
5             End If
6             If Not IsNumber(PutRatio) Then
7                 Throw "PutRatio should be a number (when StrategySwitchPoints is empty)"
8             End If
9             If Not IsNumber(CallRatio) Then
10                Throw "CallRatio should be a number (when StrategySwitchPoints is empty)"
11            End If
12            If Not IsNumber(PutStrikeOffset) Then
13                Throw "PutStrikeOffset should be a number (when StrategySwitchPoints is empty)"
14            End If
15            If Not IsNumber(CallStrikeOffset) Then
16                Throw "CallStrikeOffset should be a number (when StrategySwitchPoints is empty)"
17            End If

18            dForwardsRatio = ForwardsRatio
19            dPutRatio = PutRatio
20            dCallRatio = CallRatio
21            dPutStrikeOffset = PutStrikeOffset
22            dCallStrikeOffset = CallStrikeOffset
23        Else
              Dim i As Long
              Dim IndexToUse As Long
              Dim NumSwitchPoints As Long
              Dim SwitchOnTime As Boolean
              Dim SwitchPointsArray As Variant

24            If IsNumber(StrategySwitchPoints) Then
25                SwitchPointsArray = StrategySwitchPoints
26                Force2DArray SwitchPointsArray
27            ElseIf VarType(StrategySwitchPoints) = vbString Then
28                SwitchPointsArray = sTokeniseString(CStr(StrategySwitchPoints))
29                If LCase(SwitchPointsArray(1, 1)) = LCase("SwitchOnTime") Then
30                    SwitchOnTime = True
31                    SwitchPointsArray = sSubArray(SwitchPointsArray, 2)
32                    For i = 1 To sNRows(SwitchPointsArray)
33                        If Not IsNumeric(SwitchPointsArray(i, 1)) Then Throw SwitchPointsErrorString
34                        SwitchPointsArray(i, 1) = CDbl(SwitchPointsArray(i, 1))
35                        If i > 1 Then
36                            If SwitchPointsArray(i, 1) < SwitchPointsArray(i - 1, 1) Then
37                                Throw "StrategySwitchPoints must be listed in " & _
                                      "ascending order when first token is SwitchOnTime"
38                            End If
39                        End If
40                    Next i
41                Else
42                    For i = 1 To sNRows(SwitchPointsArray)
43                        If Not IsNumeric(SwitchPointsArray(i, 1)) Then Throw SwitchPointsErrorString
44                        SwitchPointsArray(i, 1) = CDbl(SwitchPointsArray(i, 1))
45                        If i > 1 Then
46                            If SwitchPointsArray(i, 1) > SwitchPointsArray(i - 1, 1) Then
47                                Throw "StrategySwitchPoints must be listed in descending order"
48                            End If
49                        End If
50                    Next i
51                End If
52            Else
53                Throw SwitchPointsErrorString
54            End If

55            NumSwitchPoints = sNRows(SwitchPointsArray)

56            If VarType(ForwardsRatio) <> vbString Then
57                Throw "When StrategySwitchPoints is provided, ForwardsRatio must be a comma delimited list of numbers"
58            End If

59            If VarType(PutRatio) <> vbString Then
60                Throw "When StrategySwitchPoints is provided, PutRatio must be a comma delimited list of numbers"
61            End If
62            If VarType(CallRatio) <> vbString Then
63                Throw "When StrategySwitchPoints is provided, CallRatio must be a comma delimited list of numbers"
64            End If
65            If VarType(PutStrikeOffset) <> vbString Then
66                Throw "When StrategySwitchPoints is provided, PutStrikeOffset must be a comma delimited list of numbers"
67            End If
68            If VarType(CallStrikeOffset) <> vbString Then
69                Throw "When StrategySwitchPoints is provided, CallStrikeOffset must be a comma delimited list of numbers"
70            End If

71            If SwitchOnTime Then
72                IndexToUse = 1
73                For i = 1 To NumSwitchPoints
74                    If SwitchPointsArray(i, 1) <= TimeInMonths Then
75                        IndexToUse = IndexToUse + 1
76                    Else
77                        Exit For
78                    End If
79                Next
80            Else
81                IndexToUse = 1
82                For i = 1 To NumSwitchPoints
83                    If SwitchPointsArray(i, 1) > TradeHeadroomInBillions Then
84                        IndexToUse = IndexToUse + 1
85                    Else
86                        Exit For
87                    End If
88                Next
89            End If

90            If sNRows(sTokeniseString(CStr(ForwardsRatio))) <> (NumSwitchPoints + 1) Then
91                Throw "ForwardsRatio must be a comma-delimited list of numbers with " & _
                      CStr(NumSwitchPoints + 1) & " elements"
92            End If
93            If sNRows(sTokeniseString(CStr(PutRatio))) <> (NumSwitchPoints + 1) Then
94                Throw "PutRatio must be a comma-delimited list of numbers with " & _
                      CStr(NumSwitchPoints + 1) & " elements"
95            End If
96            If sNRows(sTokeniseString(CStr(CallRatio))) <> (NumSwitchPoints + 1) Then
97                Throw "CallRatio must be a comma-delimited list of numbers with " & _
                      CStr(NumSwitchPoints + 1) & " elements"
98            End If
99            If sNRows(sTokeniseString(CStr(PutStrikeOffset))) <> (NumSwitchPoints + 1) Then
100               Throw "PutStrikeOffset must be a comma-delimited list of numbers with " & _
                      CStr(NumSwitchPoints + 1) & " elements"
101           End If
102           If sNRows(sTokeniseString(CStr(CallStrikeOffset))) <> (NumSwitchPoints + 1) Then
103               Throw "CallStrikeOffset must be a comma-delimited list of numbers with " & _
                      CStr(NumSwitchPoints + 1) & " elements"
104           End If

105           dForwardsRatio = sTokeniseString(CStr(ForwardsRatio))(IndexToUse, 1)
106           dPutRatio = sTokeniseString(CStr(PutRatio))(IndexToUse, 1)
107           dCallRatio = sTokeniseString(CStr(CallRatio))(IndexToUse, 1)
108           dPutStrikeOffset = sTokeniseString(CStr(PutStrikeOffset))(IndexToUse, 1)
109           dCallStrikeOffset = sTokeniseString(CStr(CallStrikeOffset))(IndexToUse, 1)
110       End If

111       SetOptionsStrategy = sArrayStack(dForwardsRatio, dPutRatio, dCallRatio, dPutStrikeOffset, dCallStrikeOffset)

112       Exit Function
ErrHandler:
113       Throw "#SetOptionsStrategy (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

