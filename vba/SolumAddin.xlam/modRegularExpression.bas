Attribute VB_Name = "modRegularExpression"
Option Explicit
Private Const MAXMRUsToStore = 20
Public Const MAXMRUsToShow = 9
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NegateRegEx
' Author    : Philip Swannell
' Date      : 02-May-2016
' Purpose   : Returns a regular expression that matches an arbitrary string if
'             and only if the input regex does not match it.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NegateRegEx(RegEx As String) As String
          Dim tmp As String
          'Ensure the function NegateRegEx is its own inverse
1         On Error GoTo ErrHandler
2         If Len(RegEx) >= 10 Then
3             If Left$(RegEx, 5) = "^((?!" Then
4                 If Right$(RegEx, 5) = ").)*$" Then
5                     tmp = Mid$(RegEx, 6, Len(RegEx) - 10)
6                     If VarType(sIsRegMatch(tmp, "Foo", False)) = vbBoolean Then
7                         NegateRegEx = tmp
8                         Exit Function
9                     End If
10                End If
11            End If
12        End If
13        NegateRegEx = "^((?!" & RegEx & ").)*$"
14        Exit Function
ErrHandler:
15        Throw "#NegateRegEx (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeSetValue
' Author    : Philip Swannell
' Date      : 27-Apr-2016
' Purpose   : Setting the value of a single cell to a string is harder than you'd think
'             when the string might start with special characters such as \, @ etc, and
'             Excel's behaviour is different according to whether Options > Advanced >
'             Lotus Compatibility > Transition Navigation Keys is set or not set!
' -----------------------------------------------------------------------------------------------------------------------
Sub SafeSetCellValue(c As Range, S As Variant)
          Dim SPH As clsSheetProtectionHandler
1         On Error GoTo ErrHandler
          Dim EscapeCharacter As String
          Dim origHAlign As Long
          Dim origNumberFormat As String
2         If sEquals(c.Value2, S) Then Exit Sub
3         If c.Parent.ProtectContents Then
4             Set SPH = CreateSheetProtectionHandler(c.Parent)
5         End If

6         Select Case VarType(S)
              Case vbString        'carry on to code to cope with strings...
7             Case Else
8                 c.Value = S
9                 Exit Sub
10        End Select

11        origNumberFormat = c.NumberFormat
12        origHAlign = c.HorizontalAlignment
13        If c.NumberFormat <> "@" Then c.NumberFormat = "@"
14        If origHAlign = xlHAlignCenter Then
15            EscapeCharacter = "^"
16        ElseIf origHAlign = xlHAlignRight Then
17            EscapeCharacter = """"
18        Else
19            EscapeCharacter = "'"
20        End If

21        If Not Application.TransitionNavigKeys Then
22            If Left$(S, 1) <> "\" Then        'Even though cell is number-formatted @ if the first character of the string to be entered is backslash Excel _
                                                 treats that as indicating something special, namely that the string should appear to be "smeared" across the cell.
23                c.Value = S
24            End If
25        End If
26        If c.Value <> S Then
27            c.Value = EscapeCharacter + S
28            If c.Value <> S Then
29                Throw "Attempt to set cell " + AddressND(c) + " to string value '" + S + "' failed (it took value '" & CStr(c.Value) & "' instead)"
30            End If
31        End If
32        If c.NumberFormat <> origNumberFormat Then c.NumberFormat = origNumberFormat
33        If c.HorizontalAlignment <> origHAlign Then c.HorizontalAlignment = origHAlign
34        Exit Sub
ErrHandler:
35        Throw "#SafeSetCellValue (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowRegularExpressionDialog
' Author    : Philip Swannell
' Date      : 01-May-2016
' Purpose   : Show a dialog allowing the user to construct a Regular expression that might
'             be useful to search a set of strings.
'             Dialog design follows Excel's "Custom AutoFilter" dialog under which appears when a range
'             has filters added (Ribbon Data > Filter) in-cell dropdown > Text Filters > Custom Filter...
'       Note: if the first argument is passed as a Range this method will set its value to the return from the dialog
'       If WithMRU is TRUE then we see a command bar in which:
'          the first options are "recently used filters" as long as those filters continue to be relevant
'          to the DataToFilter i.e  match some but not all of DataToFilter
'          If there are fewer than 9 such filters then the subsequent filters are the "most frequently
'          appearing" strings in DataToFilter.
'          Next there may be a "Pick from list..." item in the command bar (suppressed if other elements of
'          the command bar cover all distict elements in DataToFilter).
'          Finally there is an "Advanced filtering..." element that puts up a custom dialog allowing the user to
'          construct a regular expreesion (of quite simple form)
' -----------------------------------------------------------------------------------------------------------------------
Function ShowRegularExpressionDialog(Optional ByVal InitialRegularExpression As Variant, _
          Optional AttributeName As String, _
          Optional ByVal DataToFilter, _
          Optional AnchorObject As Object, _
          Optional Title As String = "Filter Rows", _
          Optional ActionText As String = "Show rows where:", _
          Optional WithMRU As Boolean, _
          Optional RegKey As String, _
          Optional AddReturnToMRU As Boolean = True, _
          Optional RangeForCopyFiltered As Range) As Variant

          Dim CellToSet As Range
          Dim chAdvancedFilteringNum As Long
          Dim chCopyFiltered As String
          Dim chCopyFilteredNum As Long
          Dim chClearRecent As Variant
          Dim FidClearRecent As Variant
          Dim ClearRecentNum As Long
          Dim ChooseVector As Variant
          Dim CustomChoice As Long
          Dim i As Long
          Dim isFromKeyboard As Boolean
          Dim PMChoice As Long
          Dim POChoice As Long
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim theForm As frmRegularExpression

1         On Error GoTo ErrHandler

2         Set theForm = New frmRegularExpression
3         isFromKeyboard = (sElapsedTime() - SafeMax(LastShiftF10Time, LastAltBacktickTime)) < 0.5

4         Set SUH = CreateScreenUpdateHandler(True)

5         If TypeName(InitialRegularExpression) = "Range" Then
6             Set CellToSet = InitialRegularExpression
7             InitialRegularExpression = CStr(InitialRegularExpression)
8             If AnchorObject Is Nothing Then
9                 CreatePositionInstructions AnchorObject, CellToSet.Offset(1), 0, 5
10            End If
11        End If

12        If Not IsMissing(DataToFilter) Then
13            If TypeName(DataToFilter) = "Range" Then
                  Dim c As Range, R As Range
                  Dim DataToFilter2 As Variant
                  Dim UseTextProperty As Boolean, UseCStr As Boolean, UseThisFormat As String
14                Set R = DataToFilter
15                TestToUseTextProperty R, UseTextProperty, UseCStr, UseThisFormat
16                If UseTextProperty Then
17                    DataToFilter2 = DataToFilter.Value2
18                    Force2DArray DataToFilter2
19                    i = 0
20                    For Each c In DataToFilter.Cells
21                        i = i + 1
22                        If VarType(c.Value) <> vbString Then
23                            DataToFilter2(i, 1) = c.text
24                        End If
25                    Next c
26                    DataToFilter = DataToFilter2
27                ElseIf UseCStr Then
28                    DataToFilter = DataToFilter.Value2
29                    For i = 1 To sNRows(DataToFilter)
30                        DataToFilter(i, 1) = CStr(DataToFilter(i, 1))
31                    Next
32                Else
33                    DataToFilter = DataToFilter.Value2
34                    For i = 1 To sNRows(DataToFilter)
35                        DataToFilter(i, 1) = Format$(DataToFilter(i, 1), UseThisFormat)
36                    Next
37                End If

38            Else
39                ChooseVector = sArrayIsText(DataToFilter)
40                If sColumnOr(ChooseVector)(1, 1) Then
41                    DataToFilter = sMChoose(DataToFilter, ChooseVector)
42                Else
43                    DataToFilter = CreateMissing()
44                End If
45            End If
46        End If

47        If WithMRU Then
              Dim EnableFlags
              Dim FaceIDs
              Dim Filters
              Dim Res
              Dim TheChoices
48            GetMRUFilters RegKey, Filters, TheChoices, FaceIDs, EnableFlags, DataToFilter
49            If sNRows(Filters) > 0 Then
50                chClearRecent = "Clear &History"
51                FidClearRecent = 358
52            Else
53                chClearRecent = CreateMissing()
54                FidClearRecent = CreateMissing()
55            End If

56            If sNRows(Filters) < MAXMRUsToShow Then
57                If Not IsMissing(DataToFilter) Then
                      Dim TopTen
                      'Use of sCountDistinctItems ensures that strings which appear many times in the array DataToFilter are at the top of the list of suggested filters
58                    TopTen = sCountDistinctItems(DataToFilter)
59                    If sNRows(TopTen) > 10 Then
60                        TopTen = sSubArray(TopTen, 1, 1, 10, 1)
61                    Else
62                        TopTen = sSubArray(TopTen, 1, 1, , 1)
63                    End If
64                    ChooseVector = sArrayIsText(sMatch(sRegExpFromLiteral(TopTen), Filters))
65                    If sColumnOr(ChooseVector)(1, 1) Then
66                        TopTen = sMChoose(TopTen, ChooseVector)
67                        If sNRows(TopTen) > MAXMRUsToShow - sNRows(TheChoices) Then
68                            TopTen = sSubArray(TopTen, 1, 1, MAXMRUsToShow - sNRows(TheChoices))
69                        End If
                          Dim TopTenAbbreviated
70                        TopTenAbbreviated = TopTen
71                        Force2DArray TopTenAbbreviated
72                        For i = 1 To sNRows(TopTenAbbreviated)
73                            TopTenAbbreviated(i, 1) = AbbreviateForCommandBar(CStr(TopTenAbbreviated(i, 1)), True)
74                        Next i
75                        TheChoices = sArrayStack(TheChoices, TopTenAbbreviated)
76                        Filters = sArrayStack(Filters, sRegExpFromLiteral(TopTen))
77                        FaceIDs = sArrayStack(FaceIDs, sReshape(0, sNRows(TopTen), 1))
78                        EnableFlags = sArrayStack(EnableFlags, sReshape(True, sNRows(TopTen), 1))
79                    End If
80                End If
81            End If
82            If Not IsMissing(DataToFilter) Then
83                DataToFilter = sRemoveDuplicates(DataToFilter, True)
84            End If
85            If Not IsMissing(Filters) Then
86                If sNRows(Filters) > MAXMRUsToShow Then
87                    Filters = sSubArray(Filters, 1, 1, MAXMRUsToShow)
88                    TheChoices = sSubArray(TheChoices, 1, 1, MAXMRUsToShow)
89                    FaceIDs = sSubArray(FaceIDs, 1, 1, MAXMRUsToShow)
90                    EnableFlags = sSubArray(EnableFlags, 1, 1, MAXMRUsToShow)
91                End If

                  Dim ShowPickFromList As Boolean
92                If Not IsMissing(DataToFilter) Then
93                    If sNRows(DataToFilter) > sNRows(TheChoices) Then
94                        ShowPickFromList = True
95                    ElseIf sNRows(sCompareTwoArrays(DataToFilter, sRegExpFromLiteral(Filters, True), "In1AndNotIn2")) > 1 Then
96                        ShowPickFromList = True
97                    End If
98                End If

99                If ShowPickFromList Then
100                   TheChoices = sArrayStack(TheChoices, "--&Pick one...", "Pick &Multiple...", "--&Advanced filtering...")
101                   FaceIDs = sArrayStack(FaceIDs, 447, 448, 502)
102                   EnableFlags = sArrayStack(EnableFlags, True, True, True)
103                   POChoice = sNRows(TheChoices) - 2
104                   PMChoice = sNRows(TheChoices) - 1
105                   CustomChoice = sNRows(TheChoices)
106                   chAdvancedFilteringNum = sNRows(TheChoices)
107               Else
108                   TheChoices = sArrayStack(TheChoices, "--&Advanced filtering...")
109                   chAdvancedFilteringNum = sNRows(TheChoices)
110                   FaceIDs = sArrayStack(FaceIDs, 502)
111                   EnableFlags = sArrayStack(EnableFlags, True)
112                   POChoice = -1
113                   PMChoice = -2
114                   CustomChoice = sNRows(TheChoices)
115               End If

116               If Not RangeForCopyFiltered Is Nothing Then
117                   chCopyFiltered = "&Copy visible rows"
118                   TheChoices = sArrayStack(TheChoices, "--" & chCopyFiltered)
119                   FaceIDs = sArrayStack(FaceIDs, 19)
120                   EnableFlags = sArrayStack(EnableFlags, True)
121                   chCopyFilteredNum = sNRows(TheChoices)
122               End If

123               If Not IsMissing(chClearRecent) Then
124                   TheChoices = sArrayStack(TheChoices, "--" & chClearRecent)
125                   FaceIDs = sArrayStack(FaceIDs, FidClearRecent)
126                   EnableFlags = sArrayStack(EnableFlags, True)
127                   ClearRecentNum = sNRows(TheChoices)
128               End If

129               Res = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , IIf(isFromKeyboard, AnchorObject, Nothing), True)
130               If Res = 0 Then
131                   ShowRegularExpressionDialog = "#User Cancel!"
132                   Exit Function
133               ElseIf Res = chCopyFilteredNum Then
134                   CopyVisibleRows RangeForCopyFiltered
135                   Exit Function
136               ElseIf Res = POChoice Then
137                   ShowRegularExpressionDialog = ShowSingleChoiceDialog(DataToFilter, , , , , "Choose " + AttributeName, "Search")
138                   If IsEmpty(ShowRegularExpressionDialog) Then
139                       ShowRegularExpressionDialog = "#User Cancel!"
140                   Else
141                       ShowRegularExpressionDialog = sRegExpFromLiteral(ShowRegularExpressionDialog)
142                   End If
143                   GoTo EarlyExit
144               ElseIf Res = PMChoice Then
                      Dim InitialChoices As Variant
145                   If Len(CStr(InitialRegularExpression)) > 0 Then
146                       InitialChoices = sTokeniseString(CStr(InitialRegularExpression), "|")
147                       InitialChoices = sArrayLeft(InitialChoices, -1)
148                       InitialChoices = sArrayRight(InitialChoices, -1)
149                       InitialChoices = sRegExpFromLiteral(InitialChoices, True)
150                   End If
151                   ShowRegularExpressionDialog = ShowMultipleChoiceDialog(DataToFilter, InitialChoices, "Choose " + AttributeName, , , , , , False)
152                   If Not sArraysIdentical(ShowRegularExpressionDialog, "#User Cancel!") Then
153                       ShowRegularExpressionDialog = sConcatenateStrings(sArrayConcatenate("^", sRegExpFromLiteral(ShowRegularExpressionDialog), "$"), "|")
154                   End If
155                   GoTo EarlyExit
156               ElseIf Res = ClearRecentNum Then
157                   RemoveFiltersFromMRU RegKey, True
158                   ShowRegularExpressionDialog = "#User Cancel!"
159                   GoTo EarlyExit
160               ElseIf Res <> chAdvancedFilteringNum Then
161                   ShowRegularExpressionDialog = Filters(Res, 1)
162                   GoTo EarlyExit
163               End If
164           End If
165       End If

166       If Not IsMissing(DataToFilter) Then
167           Force2DArray DataToFilter
168       End If

169       theForm.Initialise Title, ActionText, AttributeName, DataToFilter, RegularExpressionToArray(CStr(InitialRegularExpression))
170       SetFormPosition theForm, AnchorObject
171       theForm.Show

172       If sIsErrorString(theForm.ReturnArray) Then
173           ShowRegularExpressionDialog = theForm.ReturnArray
174       Else
175           Res = ArrayToRegularExpression(theForm.ReturnArray)
176           If AddReturnToMRU Then
177               AddFilterToMRU RegKey, CStr(Res)
178           End If
179           ShowRegularExpressionDialog = Res
180       End If

181       Set theForm = Nothing

EarlyExit:
182       If ShowRegularExpressionDialog <> "#User Cancel!" Then
183           If Not CellToSet Is Nothing Then
184               Set SPH = CreateSheetProtectionHandler(CellToSet.Parent)
185               SafeSetCellValue CellToSet, CStr(ShowRegularExpressionDialog)
186           End If
187       End If

188       Exit Function

ErrHandler:
189       Throw "#ShowRegularExpressionDialog(line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ArrayToRegularExpression
' Author    : Philip Swannell
' Date      : 01-May-2016
' Purpose   : Takes the array returned by the dialog and translates to a regular expression,
'             example return from the dialog:
'------------------------------------
'   |  1      2                 3
'------------------------------------
'1  |  ''     'contains'        'XX'
'2  |  'AND'  'does not equal'  'XXY'
'====================================
' -----------------------------------------------------------------------------------------------------------------------
Private Function ArrayToRegularExpression(Optional ByVal TheArray As Variant)

1         On Error GoTo ErrHandler

          Dim i As Long
          Dim RegExParts As Variant
          Dim ResultA As String
          Dim ResultB As String
          Dim ThisPart As Variant
          Dim TmpArray As Variant
          Dim TmpString As String

          'Deal with the "in" and "not in" cases first - of the elements of the second column, only the first one can be an "in" or "not in"
2         If TheArray(1, 2) = "in" Or TheArray(1, 2) = "not in" Then
3             TmpArray = sTokeniseString(CStr(TheArray(1, 3)))
4             TmpArray = sRegExpFromLiteral(TmpArray, False)
5             TmpString = sConcatenateStrings(sArrayConcatenate("^", TmpArray, "$"), "|")
6             If TheArray(1, 2) = "not in" Then
7                 TmpString = NegateRegEx(TmpString)
8             End If
9             ArrayToRegularExpression = TmpString
10            Exit Function
11        End If

12        RegExParts = sReshape(vbNullString, sNRows(TheArray), 1)

          'Special case of achieving AND with lookarounds. Likely to yield a shorter and more understandable regular expression than the triple use of NegateRegEx
13        If sNRows(TheArray) = 2 Then
14            If TheArray(2, 1) = "AND" Then
15                For i = 1 To sNRows(TheArray)
16                    Select Case CStr(TheArray(i, 2))
                          Case "does not begin with"
17                            ThisPart = "(?!^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & ")"
18                        Case "does not end with"
19                            ThisPart = "(?!.*" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$)"
20                        Case "does not contain"
21                            ThisPart = "(?!.*" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & ")"
22                        Case "does not equal"
23                            ThisPart = "(?!^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$)"
24                        Case "begins with"
25                            ThisPart = "(?=^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & ")"
26                        Case "ends with"
27                            ThisPart = "(?=.*" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$)"
28                        Case "contains"
29                            ThisPart = "(?=.*" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & ")"
30                        Case "equals"
31                            ThisPart = "(?=^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$)"
32                        Case Else
33                            Throw "unexpected error"
34                    End Select
35                    RegExParts(i, 1) = ThisPart
36                Next i
37                ResultA = "^" & RegExParts(1, 1) & RegExParts(2, 1) & ".*$"
38            End If
39        End If

          'Next the other cases
40        For i = 1 To sNRows(TheArray)
41            Select Case CStr(TheArray(i, 2))
                  Case "begins with"
42                    ThisPart = "^" & sRegExpFromLiteral(CStr(TheArray(i, 3)))
43                Case "ends with"
44                    ThisPart = sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$"
45                Case "does not contain"
46                    ThisPart = NegateRegEx(sRegExpFromLiteral(CStr(TheArray(i, 3))))
47                Case "equals"
48                    ThisPart = "^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$"
49                Case "does not equal"
50                    ThisPart = NegateRegEx("^" & sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$")
51                Case "does not begin with"
52                    ThisPart = NegateRegEx("^" & sRegExpFromLiteral(CStr(TheArray(i, 3))))
53                Case "does not end with"
54                    ThisPart = NegateRegEx(sRegExpFromLiteral(CStr(TheArray(i, 3))) & "$")
55                Case "contains"
56                    ThisPart = sRegExpFromLiteral(CStr(TheArray(i, 3)))
57                Case Else
58                    Throw "unexpected error"
59            End Select
60            RegExParts(i, 1) = ThisPart
61        Next i

62        If sNRows(TheArray) = 1 Then
63            ResultB = RegExParts(1, 1)
64        ElseIf TheArray(2, 1) = vbNullString Then
65            ResultB = RegExParts(1, 1)
66        ElseIf TheArray(2, 1) = "OR" Then
67            ResultB = RegExParts(1, 1) & "|" & RegExParts(2, 1)
68        ElseIf TheArray(2, 1) = "AND" Then
69            ResultB = NegateRegEx(NegateRegEx(CStr(RegExParts(1, 1))) & "|" & NegateRegEx(CStr(RegExParts(2, 1))))        'A and B  = not (not A or not B)
70        End If

          'Take the shorter result
71        If Len(ResultA) = 0 Then
72            ArrayToRegularExpression = ResultB
73        ElseIf Len(ResultA) < Len(ResultB) Then
74            ArrayToRegularExpression = ResultA
75        Else
76            ArrayToRegularExpression = ResultB
77        End If

78        Exit Function
ErrHandler:
79        Throw "#ArrayToRegularExpression(line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RegularExpressionToArray
' Author    : Philip Swannell
' Date      : 03-May-2016
' Purpose   : The inverse of ArrayToRegularExpression, i.e. attempts to find ResultArray
'             with three columns and one or two rows such that ArrayToRegularExpression(ResultArray) = InputRegularExpression
'             If the attempt is unsuccessful then returns Empty.
'             Note that the function cannot return an incorrect non-empty value, thanks to the line:
'             If ArrayToRegularExpression(CandidateArray) = InputRegularExpression Then...
' -----------------------------------------------------------------------------------------------------------------------
Private Function RegularExpressionToArray(InputRegularExpression As String)
          Dim CandidateArray As Variant
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim LeftPart As String
          Dim MiddlePart As String
          Dim Operator1 As String
          Dim Operator2 As String
          Dim operatorBoolean As String
          Dim RegExForTesting As String
          Dim ReplaceXWith As String
          Dim ReplaceYWith As String
          Dim RightPart As String
          Dim TemplateArray() As Variant
          Dim TemplateRegularExpression As String

          'Assume we need just one row of the dialog, and loop through the possibilities for the elements of TemplateArray
1         On Error GoTo ErrHandler
2         ReDim TemplateArray(1 To 1, 1 To 3)
3         TemplateArray(1, 3) = "X": TemplateArray(1, 1) = vbNullString
4         For i = 1 To 8
5             Operator1 = Choose(i, "contains", "equals", "does not equal", "begins with", "does not begin with", "ends with", "does not end with", "does not contain")
6             TemplateArray(1, 2) = Operator1
7             TemplateRegularExpression = ArrayToRegularExpression(TemplateArray)
8             LeftPart = sStringBetweenStrings(TemplateRegularExpression, , "X")
9             RightPart = sStringBetweenStrings(TemplateRegularExpression, "X")
10            RegExForTesting = sRegExpFromLiteral(LeftPart) + ".*" + sRegExpFromLiteral(RightPart)
11            If sIsRegMatch(RegExForTesting, InputRegularExpression) Then
12                ReplaceXWith = sStringBetweenStrings(InputRegularExpression, LeftPart, RightPart)
13                ReplaceXWith = sRegExpFromLiteral(ReplaceXWith, True)
14                CandidateArray = TemplateArray
15                CandidateArray(1, 3) = ReplaceXWith
16                If ArrayToRegularExpression(CandidateArray) = InputRegularExpression Then
17                    RegularExpressionToArray = CandidateArray
18                    Exit Function
19                End If
20            End If
21        Next i

          'Assume we need two rows of the dialog, and loop through the 128 possibilities for the elements of TemplateArray
22        ReDim TemplateArray(1 To 2, 1 To 3)
23        TemplateArray(1, 3) = "X": TemplateArray(2, 3) = "Y": TemplateArray(1, 1) = vbNullString

24        For i = 1 To 8
25            Operator1 = Choose(i, "contains", "equals", "does not equal", "begins with", "does not begin with", "ends with", "does not end with", "does not contain")
26            For j = 1 To 8
27                Operator2 = Choose(j, "contains", "equals", "does not equal", "begins with", "does not begin with", "ends with", "does not end with", "does not contain")
28                For k = 1 To 2
29                    operatorBoolean = Choose(k, "OR", "AND")
30                    TemplateArray(2, 1) = operatorBoolean: TemplateArray(1, 2) = Operator1: TemplateArray(2, 2) = Operator2
31                    TemplateRegularExpression = ArrayToRegularExpression(TemplateArray)
32                    LeftPart = sStringBetweenStrings(TemplateRegularExpression, , "X")
33                    MiddlePart = sStringBetweenStrings(TemplateRegularExpression, "X", "Y")
34                    RightPart = sStringBetweenStrings(TemplateRegularExpression, "Y")
35                    RegExForTesting = sRegExpFromLiteral(LeftPart) + ".*" + sRegExpFromLiteral(MiddlePart) + ".*" + sRegExpFromLiteral(RightPart)
36                    If sIsRegMatch(RegExForTesting, InputRegularExpression) Then
37                        ReplaceXWith = sStringBetweenStrings(InputRegularExpression, LeftPart, MiddlePart)
38                        ReplaceYWith = sStringBetweenStrings(InputRegularExpression, MiddlePart, RightPart)
39                        ReplaceXWith = sRegExpFromLiteral(ReplaceXWith, True)
40                        ReplaceYWith = sRegExpFromLiteral(ReplaceYWith, True)
41                        CandidateArray = TemplateArray
42                        CandidateArray(1, 3) = ReplaceXWith
43                        CandidateArray(2, 3) = ReplaceYWith

44                        If ArrayToRegularExpression(CandidateArray) = InputRegularExpression Then
45                            RegularExpressionToArray = CandidateArray
46                            Exit Function
47                        End If
48                    End If
49                Next k
50            Next j
51        Next i

          'Finally try the "in" and "not in" cases
          Dim TopRightElement As String
52        If Len(InputRegularExpression) > 8 Then
53            TopRightElement = Mid$(InputRegularExpression, 2, Len(InputRegularExpression) - 2)
54            TopRightElement = sConcatenateStrings(sRegExpFromLiteral(sTokeniseString(TopRightElement, "$|^"), True))
55            CandidateArray = sArrayRange(vbNullString, "in", TopRightElement)
56            If ArrayToRegularExpression(CandidateArray) = InputRegularExpression Then
57                RegularExpressionToArray = CandidateArray
58                Exit Function
59            End If
60        End If

61        If Len(InputRegularExpression) > 15 Then
62            If Left$(InputRegularExpression, 6) = "^((?!^" Then
63                If Right$(InputRegularExpression, 6) = "$).)*$" Then
64                    TopRightElement = Mid$(InputRegularExpression, 7, Len(InputRegularExpression) - 12)
65                    TopRightElement = sConcatenateStrings(sRegExpFromLiteral(sTokeniseString(TopRightElement, "$|^"), True))
66                    CandidateArray = sArrayRange(vbNullString, "not in", TopRightElement)
67                    If ArrayToRegularExpression(CandidateArray) = InputRegularExpression Then
68                        RegularExpressionToArray = CandidateArray
69                        Exit Function
70                    End If
71                End If
72            End If
73        End If

74        Exit Function
ErrHandler:
75        Throw "#RegularExpressionToArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddFilterToMRU
' Author    : Philip Swannell
' Date      : 28-Apr-2016
' Purpose   : Save a filter to the registry
' -----------------------------------------------------------------------------------------------------------------------
Sub AddFilterToMRU(RegKey As String, filter As String)
          Dim MRUFilters As Variant

1         On Error GoTo ErrHandler
2         MRUFilters = GetSetting(gAddinName, "Filters", RegKey + "MRU", "Not found")
3         If MRUFilters <> "Not found" Then
4             MRUFilters = sArrayStack(filter, sParseArrayString(CStr(MRUFilters)))
5             MRUFilters = sRemoveDuplicates(MRUFilters, False, False)
6             If sNRows(MRUFilters) > MAXMRUsToStore Then
7                 MRUFilters = sSubArray(MRUFilters, 1, 1, MAXMRUsToStore)
8             End If
9         Else
10            MRUFilters = filter
11        End If
12        SaveSetting gAddinName, "Filters", RegKey + "MRU", sMakeArrayString(MRUFilters)

13        Exit Sub
ErrHandler:
14        Throw "#AddFilterToMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveFiltersFromMRU
' Author    : Philip Swannell
' Date      : 25-Nov-2016
' Purpose   : Remove one or more filters from the registry. If Filters is Boolean True then all filters are removed
' -----------------------------------------------------------------------------------------------------------------------
Sub RemoveFiltersFromMRU(RegKey As String, Filters As Variant)
          Dim EnableFlags
          Dim ExistingFilters
          Dim FaceIDs
          Dim NewFilters
          Dim TheChoices
1         On Error GoTo ErrHandler

2         GetMRUFilters RegKey, ExistingFilters, TheChoices, FaceIDs, EnableFlags
3         If Not IsMissing(ExistingFilters) Then
4             If VarType(Filters) = vbBoolean Then
5                 If Filters Then
6                     DeleteSetting gAddinName, "Filters", RegKey + "MRU"
7                     Exit Sub
8                 End If
9             End If
              
10            NewFilters = sCompareTwoArrays(ExistingFilters, Filters, "In1AndNotIn2")
11            If sNRows(NewFilters) = 1 Then        'the header only
12                DeleteSetting gAddinName, "Filters", RegKey + "MRU"
13            Else
14                NewFilters = sSubArray(NewFilters, 2)
15                SaveSetting gAddinName, "Filters", RegKey + "MRU", sMakeArrayString(NewFilters)
16            End If
17        End If

18        Exit Sub
ErrHandler:
19        Throw "#RemoveFiltersFromMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetMRUFilters
' Author    : Philip Swannell
' Date      : 28-Apr-2016
' Purpose   : Read from the registry the filters associated with a key.
'             If DataToFilter is passed it should be a single column array, preferably with no
'             repeats. The filters returned by this method are then checked for "filtering power" against
'             DataToFilter, where filtering power means "selects some but not all".
'             Note also that the filters held in the Registry are amended to remove those that have no filtering power!
' -----------------------------------------------------------------------------------------------------------------------
Sub GetMRUFilters(RegKey As String, ByRef Filters, ByRef TheChoices, ByRef FaceIDs, ByRef EnableFlags, Optional DataToFilter As Variant)
          Dim ChooseVector As Variant
          Dim FiltersAbbreviated As Variant
          Dim i As Long
          Dim SomeBad As Boolean
          Dim SomeGood As Boolean
          Dim TRUEFALSEArray As Variant

1         On Error GoTo ErrHandler
2         Filters = GetSetting(gAddinName, "Filters", RegKey + "MRU", "Not found")
3         If Filters = "Not found" Then
4             Filters = CreateMissing(): TheChoices = CreateMissing(): FaceIDs = CreateMissing(): EnableFlags = CreateMissing()
5             Exit Sub
6         End If

7         Filters = sParseArrayString(CStr(Filters))
8         ChooseVector = sReshape(False, sNRows(Filters), 1)

9         Force2DArray Filters

10        If Not IsMissing(DataToFilter) Then
11            For i = 1 To sNRows(Filters)
12                TRUEFALSEArray = sIsRegMatch(CStr(Filters(i, 1)), DataToFilter, False)
13                If IsNumber(sMatch(True, TRUEFALSEArray)) And IsNumber(sMatch(False, TRUEFALSEArray)) Then
14                    ChooseVector(i, 1) = True
15                    SomeGood = True
16                Else
17                    SomeBad = True
18                End If
19            Next i
20            If SomeBad Then
21                RemoveFiltersFromMRU RegKey, sMChoose(Filters, sArrayNot(ChooseVector))
22            End If
23            If Not SomeGood Then
24                Filters = CreateMissing(): TheChoices = CreateMissing(): FaceIDs = CreateMissing(): EnableFlags = CreateMissing()
25                Exit Sub
26            Else
27                Filters = sMChoose(Filters, ChooseVector)
28                Force2DArray Filters
29            End If
30        End If

31        If sNRows(Filters) > MAXMRUsToStore Then
32            Filters = sSubArray(Filters, 1, 1, MAXMRUsToStore)
33        End If

34        FiltersAbbreviated = Filters
35        For i = 1 To sNRows(FiltersAbbreviated)
36            FiltersAbbreviated(i, 1) = AbbreviateForCommandBar(sRegExpFromLiteral(CStr(FiltersAbbreviated(i, 1)), True))
37        Next i

38        TheChoices = FiltersAbbreviated
39        FaceIDs = sReshape(0, sNRows(TheChoices), 1)
40        For i = 1 To SafeMin(9, sNRows(TheChoices))
41            FaceIDs(i, 1) = 70 + i Mod 10
42        Next i

43        EnableFlags = sReshape(True, sNRows(TheChoices), 1)

44        Exit Sub
ErrHandler:
45        Throw "#GetMRUFilters (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AllTheRegExs
' Author    : Philip Swannell
' Date      : 14-May-2016
' Purpose   : Subroutine of RunTest. Returns an array showing all the possible returns from ArrayToRegularExpression
'             assuming that the two strings are "X" and "Y" and that the Array processed by ATRE
'             has two rows. The return has 128 rows (8 x 8 x 2) with an example row of the array being:
'             'contains' 'X' 'OR' 'begins with' 'Y' 'X|^Y'
' -----------------------------------------------------------------------------------------------------------------------
Private Function AllTheRegExs()
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim Operator1 As String
          Dim Operator2 As String
          Dim operatorBoolean As String
          Dim Result As Variant
          Dim TemplateArray As Variant
          Dim TemplateRegularExpression As String
          Dim w As Long

1         On Error GoTo ErrHandler
2         TemplateArray = sReshape(vbNullString, 2, 3)
3         TemplateArray(1, 3) = "X"
4         TemplateArray(2, 3) = "Y"
5         Result = sReshape(vbNullString, 128, 6)
6         w = 0
7         For k = 1 To 2
8             operatorBoolean = Choose(k, "OR", "AND")
9             For i = 1 To 8
10                Operator1 = Choose(i, "contains", "equals", "does not equal", "begins with", "does not begin with", "ends with", "does not end with", "does not contain")
11                For j = 1 To 8
12                    Operator2 = Choose(j, "contains", "equals", "does not equal", "begins with", "does not begin with", "ends with", "does not end with", "does not contain")
13                    TemplateArray(2, 1) = operatorBoolean: TemplateArray(1, 2) = Operator1: TemplateArray(2, 2) = Operator2
14                    TemplateRegularExpression = ArrayToRegularExpression(TemplateArray)
15                    w = w + 1
16                    Result(w, 1) = Operator1
17                    Result(w, 2) = "X"
18                    Result(w, 3) = operatorBoolean
19                    Result(w, 4) = Operator2
20                    Result(w, 5) = "Y"
21                    Result(w, 6) = TemplateRegularExpression

22                Next j
23            Next i
24        Next k
25        AllTheRegExs = Result

26        Exit Function
ErrHandler:
27        AllTheRegExs = "#AllTheRegExs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StringsToTest
' Author    : Philip Swannell
' Date      : 14-May-2016
' Purpose   : Sub routine of RunTest. Return a list of strings that lists all possible
'             combinations  of the characters W,X,Y,Z of length up to 5, so return has
'             4^5 + 4^4 + 4^3 +4^2 + 4 = 1364 elements
' -----------------------------------------------------------------------------------------------------------------------
Private Function StringsToTest()

          Dim i As Long
          Static Res As Variant

1         On Error GoTo ErrHandler
2         If IsEmpty(Res) Then
3             Res = sArrayMakeText(sIntegers(99999))
4             For i = 1 To sNRows(Res)
5                 Res(i, 1) = Replace(Res(i, 1), "0", "W")
6                 Res(i, 1) = Replace(Res(i, 1), "1", "X")
7                 Res(i, 1) = Replace(Res(i, 1), "2", "Y")
8                 Res(i, 1) = Replace(Res(i, 1), "3", "Z")
9                 Res(i, 1) = Replace(Res(i, 1), "4", "W")
10                Res(i, 1) = Replace(Res(i, 1), "5", "X")
11                Res(i, 1) = Replace(Res(i, 1), "6", "Y")
12                Res(i, 1) = Replace(Res(i, 1), "7", "Z")
13                Res(i, 1) = Replace(Res(i, 1), "8", "W")
14                Res(i, 1) = Replace(Res(i, 1), "9", "X")
15            Next i
16            Res = sRemoveDuplicates(Res, True)
17        End If

18        StringsToTest = Res
19        Exit Function
ErrHandler:
20        Throw "#StringsToTest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MatchStringToRules
' Author    : Philip Swannell
' Date      : 14-May-2016
' Purpose   : Sub-routine of RunTest. Instead of matching a string against a Regular Expression
'             apply the rules that that regular expression purports to represent, according to method
'             ArrayToRegularExpression.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MatchStringToRules(StringToMatch As String, Operator1 As String, MatchString1 As String, operatorBoolean As String, Operator2 As String, MatchString2 As String)
          Dim Res1 As Boolean
          Dim Res2 As Boolean

1         On Error GoTo ErrHandler
2         Select Case Operator1
              Case "begins with"
3                 Res1 = LCase$(Left$(StringToMatch, Len(MatchString1))) = LCase$(MatchString1)
4             Case "ends with"
5                 Res1 = LCase$(Right$(StringToMatch, Len(MatchString1))) = LCase$(MatchString1)
6             Case "does not contain"
7                 Res1 = InStr(LCase$(StringToMatch), LCase$(MatchString1)) = 0
8             Case "equals"
9                 Res1 = LCase$(StringToMatch) = LCase$(MatchString1)
10            Case "does not equal"
11                Res1 = LCase$(StringToMatch) <> LCase$(MatchString1)
12            Case "does not begin with"
13                Res1 = LCase$(Left$(StringToMatch, Len(MatchString1))) <> LCase$(MatchString1)
14            Case "does not end with"
15                Res1 = LCase$(Right$(StringToMatch, Len(MatchString1))) <> LCase$(MatchString1)
16            Case "contains"
17                Res1 = InStr(LCase$(StringToMatch), LCase$(MatchString1)) <> 0
18            Case Else
19                Throw "Unrecognised operator1"
20        End Select

21        Select Case Operator2
              Case "begins with"
22                Res2 = LCase$(Left$(StringToMatch, Len(MatchString2))) = LCase$(MatchString2)
23            Case "ends with"
24                Res2 = LCase$(Right$(StringToMatch, Len(MatchString2))) = LCase$(MatchString2)
25            Case "does not contain"
26                Res2 = InStr(LCase$(StringToMatch), LCase$(MatchString2)) = 0
27            Case "equals"
28                Res2 = LCase$(StringToMatch) = LCase$(MatchString2)
29            Case "does not equal"
30                Res2 = LCase$(StringToMatch) <> LCase$(MatchString2)
31            Case "does not begin with"
32                Res2 = LCase$(Left$(StringToMatch, Len(MatchString2))) <> LCase$(MatchString2)
33            Case "does not end with"
34                Res2 = LCase$(Right$(StringToMatch, Len(MatchString2))) <> LCase$(MatchString2)
35            Case "contains"
36                Res2 = InStr(LCase$(StringToMatch), LCase$(MatchString2)) <> 0
37            Case Else
38                Throw "Unrecognised operator2"
39        End Select

40        Select Case operatorBoolean
              Case "OR"
41                MatchStringToRules = Res1 Or Res2
42            Case "AND"
43                MatchStringToRules = Res1 And Res2
44            Case Else
45                Throw "Unrecognised OperatorBoolean"
46        End Select
47        Exit Function
ErrHandler:
48        Throw "#MatchStringToRules (line " & CStr(Erl) + "): " & Err.Description & "!"
49        Exit Function
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunTest
' Author    : Philip Swannell
' Date      : 14-May-2016
' Purpose   : Do the Regular expressions generated by ArrayToRegularExpression do what they say on the tin?
'             Test all 128 regular expressions generated by ArrayToRegularExpression
'             for all combinations of Operator1, Operator2, OperatorBoolean against
'             all 1364 elements of StringsToTest. Use method ApplyRulesToString as comparison method.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RunTest()

          Dim i As Long
          Dim j As Long
          Dim MatchString1 As String
          Dim MatchString2 As String
          Dim Message As String
          Dim NR As Long
          Dim NS As Long
          Dim Operator1 As String
          Dim Operator2 As String
          Dim operatorBoolean As String
          Dim RegArray
          Dim RegularExpression As String
          Dim Res1
          Dim Res2
          Dim StringToMatch As String
          Dim TheStrings

1         On Error GoTo ErrHandler
2         RegArray = AllTheRegExs()
3         TheStrings = StringsToTest()
4         NR = sNRows(RegArray)
5         NS = sNRows(StringsToTest)

6         For i = 1 To NR
7             Application.StatusBar = CStr(i) + "/" + CStr(NR)
8             DoEvents
9             Operator1 = RegArray(i, 1)
10            MatchString1 = RegArray(i, 2)
11            operatorBoolean = RegArray(i, 3)
12            Operator2 = RegArray(i, 4)
13            MatchString2 = RegArray(i, 5)
14            RegularExpression = RegArray(i, 6)
15            For j = 1 To NS
16                StringToMatch = TheStrings(j, 1)
17                Res1 = sIsRegMatch(RegularExpression, StringToMatch, False)
18                Res2 = MatchStringToRules(StringToMatch, Operator1, MatchString1, operatorBoolean, Operator2, MatchString2)
19                If Res1 <> Res2 Then
20                    Message = "Failure found:" + vbLf + _
                          "StringtoMatch = " + StringToMatch + vbLf + _
                          "Operator1 = " + Operator1 + vbLf + _
                          "MatchString1 = " + MatchString1 + vbLf + _
                          "operatorBoolean = " + operatorBoolean + vbLf + _
                          "Operator2 = " + Operator2 + vbLf + _
                          "MatchString2 = " + MatchString2 + vbLf + _
                          "RegularExpression = " + RegularExpression + vbLf + _
                          "RegularExpression result = " + CStr(Res1) + vbLf + _
                          "MatchStringToRules result = " + CStr(Res2)
21                    Throw Message, True
22                End If
23            Next j
24        Next i
25        MsgBoxPlus "All tests passed", vbOKOnly + vbInformation
26        Application.StatusBar = False
27        Exit Sub
ErrHandler:
28        SomethingWentWrong "#RunTest (line " & CStr(Erl) + "): " & Err.Description & "!"
29        Application.StatusBar = False
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CopyVisibleRows
' Author     : Philip Swannell
' Date       : 26-Jun-2019
' Purpose    : Copies to the clipboard those rows in ParentRange whose row height is not zero
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CopyVisibleRows(ParentRange As Range)

          Dim ChooseVector As Variant
          Dim i As Long
          Dim rngCopy As Range

1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler

3         ChooseVector = sReshape(True, ParentRange.Rows.Count, 1)

4         For i = 1 To ParentRange.Rows.Count
5             If ParentRange.Cells(i, 1).RowHeight = 0 Then
6                 ChooseVector(i, 1) = False
7             End If
8         Next

9         sFilterRange ParentRange, ChooseVector, rngCopy

10        If Not rngCopy Is Nothing Then
11            rngCopy.Copy
12        End If

13        Exit Sub
ErrHandler:
14        Throw "#CopyVisibleRows (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
