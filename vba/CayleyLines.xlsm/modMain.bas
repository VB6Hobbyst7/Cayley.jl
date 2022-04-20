Attribute VB_Name = "modMain"
Option Explicit
Private Const EditTitle = "Update Rates Notional Weights"

Sub ReleaseCleanup()
          Dim Res
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Application.ScreenUpdating = False
3         For Each ws In ThisWorkbook.Worksheets
4             ws.Visible = xlSheetVisible
5             Application.GoTo ws.Cells(1, 1)
6             ws.Protect , True, True
7             Res = ws.UsedRange.Rows.Count    'resets used range
8             ActiveWindow.DisplayGridlines = False: ActiveWindow.DisplayHeadings = False
9         Next ws
10        shComments.Visible = xlSheetHidden

11        CancelButton
12        Application.GoTo shSummary.Cells(1, 1)
13        Exit Sub
ErrHandler:
14        Throw "#ReleaseCleanup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EditNotionalWeights
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Takes user into the "Edit Notional Weights" sheet, with that sheet correctly initialised
'---------------------------------------------------------------------------------------
Sub EditNotionalWeights(BankName As String, ExistingArrayString As String, IsRates As Boolean)

          Const DefaultArrayString1 = "{""Tenor"",""EUR"",""Other"";""1Y"",0.005,0.005;""2Y"",0.007,0.007;""3Y"",0.01,0.01;""4Y"",0.0125,0.0125;""5Y"",0.016,0.016;""7Y"",0.03,0.03}"
          Const DefaultArrayString2 = "{""1Y"",0.12;""2Y"",0.17;""3Y"",0.21;""4Y"",0.24;""5Y"",0.27;""7Y"",0.32}"

          Dim SPH As clsSheetProtectionHandler

          Dim ExistingData
1         On Error GoTo ErrHandler
2         ExistingData = sParseArrayString(ExistingArrayString)

3         If sIsErrorString(ExistingData) Then
4             If IsRates Then
5                 ExistingData = sParseArrayString(DefaultArrayString1)
6             Else
7                 ExistingData = sParseArrayString(DefaultArrayString2)
8             End If
9         End If

10        Set SPH = CreateSheetProtectionHandler(shEditNotionalWeights)
11        CleanOutEditor

12        shEditNotionalWeights.Visible = xlSheetVisible
13        shSummary.Visible = xlSheetHidden
14        shAudit.Visible = xlSheetHidden

15        With shEditNotionalWeights.Cells(1, 1)
16            If IsRates Then
17                .Value = "Edit Rates Notional Weights"
18            Else
19                .Value = "Edit Fx Notional Weights"
20            End If
21            .Font.Size = 22
22            .ColumnWidth = 2
23        End With

24        With shEditNotionalWeights.Cells(5, 3)
25            .Value = "BankName"
26            .HorizontalAlignment = xlHAlignRight
27        End With
28        With shEditNotionalWeights.Cells(5, 4)
29            .Value = BankName
30            .Parent.Names.Add "BankName", .Offset(0)
31        End With

32        With shEditNotionalWeights.Cells(7, 2).Resize(snrows(ExistingData), sNCols(ExistingData))
33            .Value = ExistingData
34            AddGreyBorders .Offset(0), False
35            .Columns(1).Font.Bold = True
36            If IsRates Then
37                .Rows(1).Font.Bold = True
38            End If
39            .HorizontalAlignment = xlHAlignCenter
40            .Parent.Names.Add "TopLeftCell", .Cells(1, 1)
41            If IsRates Then
42                .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1).NumberFormat = "0.0%"
43            Else
44                .Offset(0, 1).Resize(, .Columns.Count - 1).NumberFormat = "0.0%"
45            End If
46            .Columns.AutoFit
47        End With
48        Set SPH = Nothing
49        shEditNotionalWeights.Unprotect
50        Exit Sub
ErrHandler:
51        Throw "#EditNotionalWeights (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : adr
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Wrapper to .Address
'---------------------------------------------------------------------------------------
Private Function adr(R As Range)
1         adr = Replace(R.Address, "$", "")
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateRatesNotionalWeights
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Check that the user has entered valid data
'---------------------------------------------------------------------------------------
Function ValidateRatesNotionalWeights(NotionalWeights As Range) As Boolean

          Dim AllowedLeftLabels
          Dim AllowedTopLabels
          Dim LeftLabelsRange As Range
          Dim LeftLabelsStatus As Variant
          Dim NumbersRange As Range
          Dim NumErrors As Long
          Dim STK As clsStacker
          Dim TopLabelsRange As Range

1         On Error GoTo ErrHandler
2         If NotionalWeights.Rows.Count < 2 Then Throw "NotionalWeights must have at least 2 rows"
3         If NotionalWeights.Columns.Count < 2 Then Throw "NotionalWeights must have at least 2 columns"

          Dim C As Range
4         AllowedLeftLabels = sSortedArray(sArrayStack(sArrayConcatenate(sintegers(12), "M"), sArrayConcatenate(sintegers(30), "Y")))
5         AllowedTopLabels = sSortedArray(sArrayStack(sCurrencies(False, False), "Other"))
6         Set STK = CreateStacker()

7         With NotionalWeights
8             Set TopLabelsRange = .Cells(1, 2).Resize(1, .Columns.Count - 1)
9             Set LeftLabelsRange = .Cells(2, 1).Resize(.Rows.Count - 1)
10            Set NumbersRange = .Cells(2, 2).Resize(.Rows.Count - 1, .Columns.Count - 1)
11            If .Cells(1, 1).Value <> "Tenor" Then
12                STK.StackData "Top left cell of Notional Weights (cell " + adr(.Cells(1, 1)) + ") must read 'Tenor'"
13                NumErrors = NumErrors + 1
14            End If
15        End With

16        For Each C In TopLabelsRange.Cells
17            If Not IsNumber(sMatch(CStr(C.Value), AllowedTopLabels, True)) Then
18                STK.StackData "Labels in the top row must be valid currency codes or the text 'Other', but cell " + adr(C) + " is not"
19                NumErrors = NumErrors + 1
20            End If
21        Next C

22        LeftLabelsStatus = sReshape(False, NotionalWeights.Rows.Count, 1)
23        For Each C In LeftLabelsRange.Cells
24            If Not IsNumber(sMatch(CStr(C.Value), AllowedLeftLabels, True)) Then
25                STK.StackData "Labels in the left column must indicate a number of months or years, e.g. '6M' or '5Y' but cell " + adr(C) + " does not"
26                NumErrors = NumErrors + 1
27            Else
28                LeftLabelsStatus(C.Row - NotionalWeights.Row + 1, 1) = True
29            End If
30            If LeftLabelsStatus(C.Row - NotionalWeights.Row + 1 - 1, 1) = True Then
31                If LeftLabelsStatus(C.Row - NotionalWeights.Row + 1, 1) = True Then
32                    If TenorToTime(C.Value) <= TenorToTime(C.Offset(-1).Value) Then
33                        STK.StackData "Labels in the left column must be arranged in increasing tenor, but " + adr(C) + " is out of order"
34                        NumErrors = NumErrors + 1
35                    End If
36                End If
37            End If
38        Next C

39        For Each C In NumbersRange.Cells
40            If Not IsNumber(C.Value) Then
41                STK.StackData "All notional weights must be non-negative numbers, but cell " + adr(C) + " is not"
42                NumErrors = NumErrors + 1
43            ElseIf C.Value < 0 Then
44                STK.StackData "All notional weights must be non-negative numbers, but cell " + adr(C) + " is not"
45                NumErrors = NumErrors + 1
46            End If
47            If C.Row > NotionalWeights.Row + 1 Then
48                If IsNumber(C.Value) Then
49                    If IsNumber(C.Offset(-1).Value) Then
50                        If C.Value < C.Offset(-1).Value Then
51                            STK.StackData "Notional Weights cannot decrease with maturity, but cell " + adr(C) + " does"
52                            NumErrors = NumErrors + 1
53                        End If
54                    End If
55                End If
56            End If
57        Next C

          Dim Prompt As String

58        If NumErrors > 0 Then
59            If NumErrors < 10 Then
60                Prompt = "Some of the data is not valid:" + vbLf + _
                           sConcatenateStrings(STK.Report, vbLf) + vbLf + vbLf + _
                           "Please fix those problems and try again."
61            Else
62                Prompt = "Some of the data is not valid, for example:" + vbLf + _
                           sConcatenateStrings(sSubArray(STK.Report, 1, 1, , 1), vbLf) + vbLf + vbLf + _
                           "Please fix the problems and try again."
63            End If

64            MsgBoxPlus Prompt, vbExclamation + vbOKOnly, EditTitle, , , , , 600
65            ValidateRatesNotionalWeights = False
66        Else
67            ValidateRatesNotionalWeights = True
68        End If

69        Exit Function
ErrHandler:
70        Throw "#ValidateRatesNotionalWeights (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateFxNotionalWeights
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Check that the user has entered valid data
'---------------------------------------------------------------------------------------
Function ValidateFxNotionalWeights(NotionalWeights As Range) As Boolean

          Dim AllowedLeftLabels
          Dim LeftLabelsRange As Range
          Dim LeftLabelsStatus As Variant
          Dim NumbersRange As Range
          Dim NumErrors As Long
          Dim STK As clsStacker

1         On Error GoTo ErrHandler
2         If NotionalWeights.Rows.Count < 2 Then Throw "NotionalWeights must have at least 2 rows"
3         If NotionalWeights.Columns.Count <> 2 Then Throw "NotionalWeights must have 2 columns"

          Dim C As Range
4         AllowedLeftLabels = sSortedArray(sArrayStack(sArrayConcatenate(sintegers(12), "M"), sArrayConcatenate(sintegers(30), "Y")))
5         Set STK = CreateStacker()

6         With NotionalWeights
7             Set LeftLabelsRange = .Columns(1)
8             Set NumbersRange = .Columns(2)
9         End With

10        LeftLabelsStatus = sReshape(False, NotionalWeights.Rows.Count, 1)
11        For Each C In LeftLabelsRange.Cells
12            If Not IsNumber(sMatch(CStr(C.Value), AllowedLeftLabels, True)) Then
13                STK.StackData "Labels in the left column must indicate a number of months or years, e.g. '6M' or '5Y' but cell " + adr(C) + " does not"
14                NumErrors = NumErrors + 1
15            Else
16                LeftLabelsStatus(C.Row - NotionalWeights.Row + 1, 1) = True
17            End If
18            If C.Row > NotionalWeights.Row Then
19                If LeftLabelsStatus(C.Row - NotionalWeights.Row + 1 - 1, 1) = True Then
20                    If LeftLabelsStatus(C.Row - NotionalWeights.Row + 1, 1) = True Then
21                        If TenorToTime(C.Value) <= TenorToTime(C.Offset(-1).Value) Then
22                            STK.StackData "Labels in the left column must be arranged in increasing tenor, but " + adr(C) + " is out of order"
23                            NumErrors = NumErrors + 1
24                        End If
25                    End If
26                End If
27            End If
28        Next C

29        For Each C In NumbersRange.Cells
30            If Not IsNumber(C.Value) Then
31                STK.StackData "All notional weights must be non-negative numbers, but cell " + adr(C) + " is not"
32                NumErrors = NumErrors + 1
33            ElseIf C.Value < 0 Then
34                STK.StackData "All notional weights must be non-negative numbers, but cell " + adr(C) + " is not"
35                NumErrors = NumErrors + 1
36            End If
37            If C.Row > NotionalWeights.Row Then
38                If IsNumber(C.Value) Then
39                    If IsNumber(C.Offset(-1).Value) Then
40                        If C.Value < C.Offset(-1).Value Then
41                            STK.StackData "Notional Weights cannot decrease with maturity, but cell " + adr(C) + " does"
42                            NumErrors = NumErrors + 1
43                        End If
44                    End If
45                End If
46            End If
47        Next C

          Dim Prompt As String

48        If NumErrors > 0 Then
49            If NumErrors < 10 Then
50                Prompt = "Some of the data is not valid:" + vbLf + _
                           sConcatenateStrings(STK.Report, vbLf) + vbLf + vbLf + _
                           "Please fix those problems and try again."
51            Else
52                Prompt = "Some of the data is not valid, for example:" + vbLf + _
                           sConcatenateStrings(sSubArray(STK.Report, 1, 1, , 1), vbLf) + vbLf + vbLf + _
                           "Please fix the problems and try again."
53            End If

54            MsgBoxPlus Prompt, vbExclamation + vbOKOnly, EditTitle, , , , , 600
55            ValidateFxNotionalWeights = False
56        Else
57            ValidateFxNotionalWeights = True
58        End If

59        Exit Function
ErrHandler:
60        Throw "#ValidateFxNotionalWeights (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : TenorToTime
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : flip "6M" to 0.5 etc
'---------------------------------------------------------------------------------------
Private Function TenorToTime(Tenor As String)
          Dim Res As Double
          Dim TheNumber As Double

1         On Error GoTo ErrHandler
2         TheNumber = -100
3         On Error Resume Next
4         TheNumber = CDbl(Left(CStr(Tenor), Len(CStr(Tenor)) - 1))
5         On Error GoTo ErrHandler
6         If TheNumber = -100 Then Throw "Unrecognised Tenor: " + CStr(Tenor)
7         Select Case UCase(Right(Tenor, 1))
          Case "Y"
8             Res = TheNumber
9         Case "M"
10            Res = TheNumber / 12
11        Case "W"
12            Res = TheNumber * 7 / 365.25
13        Case "D"
14            Res = TheNumber / 365.25
15        Case Else
16            Throw "Unrecognised Tenor: " + CStr(Tenor)
17        End Select
18        TenorToTime = Res
19        Exit Function
ErrHandler:
20        Throw "#TenorToTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CleanOutEditor
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : clear out everything on the sheet apart from buttons
'---------------------------------------------------------------------------------------
Sub CleanOutEditor()
          Dim N As Name
          Dim Res
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shEditNotionalWeights)

3         With shEditNotionalWeights
4             For Each N In .Names
5                 N.Delete
6             Next N
7             .UsedRange.EntireRow.Delete
8             .UsedRange.EntireColumn.Delete
9             Res = .UsedRange.Rows.Count
10        End With

11        Exit Sub
ErrHandler:
12        Throw "#CleanOutEditor (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CancelButton
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Attached to the Cancel button on the EditNotionalWeights sheet
'---------------------------------------------------------------------------------------
Sub CancelButton()
1         On Error GoTo ErrHandler
          Dim SUH As clsScreenUpdateHandler

2         Set SUH = CreateScreenUpdateHandler()
3         shSummary.Visible = xlSheetVisible
4         shAudit.Visible = xlSheetVisible
5         shEditNotionalWeights.Visible = xlSheetHidden
6         CleanOutEditor
7         shSummary.Activate
8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#CancelButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OKButton
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : Response to user clicking the OK button on the EditNotionalWeights sheet, pastes the
'             notional weights array into the appropriate cell of the Summary sheet
'---------------------------------------------------------------------------------------
Sub OKButton()
          Dim BankName As String
          Dim ColNo As Variant
          Dim IsRates As Boolean
          Dim NotionalWeights As Variant
          Dim NWRange As Range
          Dim OldNotionalWeights
          Dim Prompt
          Dim RowNo As Variant
          Dim SPH As clsSheetProtectionHandler
          Dim TargetHeader
          Dim TheHeaders As Range
          Dim TheTableNoHeaders As Range

1         On Error GoTo ErrHandler
2         Set TheHeaders = shSummary.ListObjects(1).HeaderRowRange
3         Set TheTableNoHeaders = shSummary.ListObjects(1).DataBodyRange
4         Select Case shEditNotionalWeights.Cells(1, 1)
          Case "Edit Fx Notional Weights"
5             IsRates = False
6             TargetHeader = "Fx Notional Weights"
7         Case "Edit Rates Notional Weights"
8             IsRates = True
9             TargetHeader = "Rates Notional Weights"
10        Case Else
11            Throw "Cannot determine whether to save as Rates Notional Weights or Fx Notional Weights"
12        End Select

13        ColNo = sMatch("CPTY_PARENT", sArrayTranspose(TheHeaders.Value))
14        If Not IsNumber(ColNo) Then Throw "Cannot find column headed 'CPTY_PARENT' on sheet Summary"

15        BankName = RangeFromSheet(shEditNotionalWeights, "BankName").Value

16        RowNo = sMatch(BankName, TheTableNoHeaders.Columns(ColNo))
17        If Not IsNumber(RowNo) Then Throw "Cannot find row for bank '" + BankName + "' on sheet Summary"

18        ColNo = sMatch(TargetHeader, sArrayTranspose(TheHeaders.Value))
19        If Not IsNumber(ColNo) Then Throw "Cannot find column headed '" + TargetHeader + "' on sheet Summary"

20        Set NWRange = sExpandRightDown(RangeFromSheet(shEditNotionalWeights, "TopLeftCell"))
21        If IsRates Then
22            If Not (ValidateRatesNotionalWeights(NWRange)) Then Exit Sub
23        Else
24            If Not (ValidateFxNotionalWeights(NWRange)) Then Exit Sub
25        End If

26        NotionalWeights = NWRange.Value

27        OldNotionalWeights = TheTableNoHeaders.Cells(RowNo, ColNo).Value
28        If Not IsEmpty(OldNotionalWeights) Then
29            OldNotionalWeights = sParseArrayString(CStr(OldNotionalWeights))
30        End If

31        If sArraysIdentical(OldNotionalWeights, NotionalWeights) Then
32            MsgBoxPlus "No edits have been made to Rate Notional Weights for " + BankName, vbInformation + vbOKOnly
33            CancelButton
34            Exit Sub
35        End If

36        Prompt = "Do you want to update " + TargetHeader + " for bank " + BankName + "?" + vbLf + vbLf + _
                   "Old Weights:" + vbLf + _
                   sConcatenateStrings(sJustifyArrayOfStrings(OldNotionalWeights, "Calibri", 11, vbTab), vbLf) + vbLf + vbLf + _
                   "New Weights:" + vbLf + _
                   sConcatenateStrings(sJustifyArrayOfStrings(NotionalWeights, "Calibri", 11, vbTab), vbLf)

37        If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, EditTitle, "Yes, update", "No, do nothing") <> vbOK Then Exit Sub

38        Set SPH = CreateSheetProtectionHandler(shSummary)

39        CancelButton

40        With TheTableNoHeaders.Cells(RowNo, ColNo)
41            SafeSetCellValue .Offset(0), sMakeArrayString(NotionalWeights)
42            .HorizontalAlignment = xlHAlignLeft
43        End With

44        Exit Sub
ErrHandler:
45        SomethingWentWrong "#OKButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HelpButton
' Author    : Philip Swannell
' Date      : 14-Nov-2016
' Purpose   : A "how to use this sheet" dialog
'---------------------------------------------------------------------------------------
Sub HelpButton()
          Dim Prompt

1         On Error GoTo ErrHandler
2         Prompt = "Help on Notional Weights" + vbLf + vbLf + _
                    "For each bank that uses a Notional-Based credit methodology, the lines workbook " + _
                    "holds a table of notional weights for Fx trades and a table for interest rate swaps. Line usage is " + _
                   "the trade's PV plus a function of the trade's currency and maturity." + vbLf + vbLf + _
                   "Tables are held in compressed form in the ""Fx Notional Weights"" and ""Rates Notional Weights"" columns of the Summary sheet and an editor worksheet makes it easy to edit the tables. To use it: " + vbLf + vbLf + _
                   "a) Double-click on a cell in the relevant column of the Summary sheet, and click OK in the confirmation dialog." + vbLf + _
                   "b) Edit the data. You can edit the data as necessary, adding or deleting columns or rows, but make sure that the tenors in the left column are correctly ordered and that " + _
                   "notional weights increase with tenor - this is checked at the time you save back the data." + vbLf + _
                   "c) Click the OK button to save your edits to the ""Summary"" sheet." + vbLf + vbLf + _
                   "For interest rate swaps, the column ""Other"" applies to any currency not explicitly given in the top row. Cross currency swaps use the same notional weights as Fx trades."

3         MsgBoxPlus Prompt, vbInformation + vbOKOnly, "Help on Notional Weights", , , , , 350
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#HelpButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


