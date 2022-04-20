Attribute VB_Name = "modFeed"
Option Explicit
Private mCalcCounter As Long
Public gDoFx As Boolean

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedAnchorDate
' Author    : Philip Swannell
' Date      : 20-Jun-2016
' Purpose   : Updates the Anchor date of the market...
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedAnchorDate(Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler
          Dim SPH As clsSheetProtectionHandler

2         Set SPH = CreateSheetProtectionHandler(shConfig)
3         If Live Then
4             RangeFromSheet(shConfig, "AnchorDate").Value2 = Date
5         Else
6             RangeFromSheet(shConfig, "AnchorDate").Value2 = AsOfDate
7         End If
8         Exit Sub
ErrHandler:
9         Throw "#FeedAnchorDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedAllRatesAllSheets
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Do as much updating from Bloomgberg as we can
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedAllRatesAllSheets(Live As Boolean, AsOfDate As Long, CurrencyList As Variant, DoFx As Boolean, DoSwaps As Boolean, DoBasis As Boolean, _
                          DoSwaptions As Boolean, DoCredit As Boolean, DoInflation As Boolean, DoInflationSets As Boolean)
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
          Dim i As Long

2         FeedAnchorDate Live, AsOfDate

3         ClearHiddenSheet
4         gDoFx = DoFx
5         If DoFx Then
6             FeedAllFxData Live, AsOfDate
7         End If

8         If DoCredit Then
9             FeedCreditSpreads Live, AsOfDate
10        End If
11        If DoInflation Then
12            FeedInflationSwaps Live, AsOfDate
13        End If
14        If DoInflationSets Then
15            FeedInflationIndexes Live, AsOfDate
16        End If

17        If DoSwaps Or DoBasis Or DoSwaptions Then
18            For i = 1 To sNRows(CurrencyList)
19                Set ws = ThisWorkbook.Worksheets(CurrencyList(i, 1))
20                ws.Calculate
21                If DoSwaps Then
22                    StatusBarWrap "Calling Bloomberg for " + Left(ws.Name, 3) + " swaps"
23                    FeedSwapsOnSheet ws, Live, AsOfDate
24                End If
25                If DoBasis Then
26                    StatusBarWrap "Calling Bloomberg for " + Left(ws.Name, 3) + " basis swaps"
27                    FeedBasisSwapsOnSheet ws, Live, AsOfDate
28                End If
29                If DoSwaptions Then
30                    StatusBarWrap "Calling Bloomberg for " + Left(ws.Name, 3) + " swaption vols"
31                    FeedSwaptionsOnSheet ws, Live, AsOfDate
32                End If
33            Next
34        End If

35        FeedFromHiddenSheet

36        Exit Sub
ErrHandler:
37        Throw "#FeedAllRatesAllSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedSwaptionsOnSheet
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Wrapper to FeedSwaptionVols, handles the swaption vols on one sheet - i.e. pastes the necessary calls to BDH\BDP to the Hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedSwaptionsOnSheet(sh As Worksheet, Live As Boolean, AsOfDate As Long)
          Dim VolRange As Range, QuoteType As String, Contributor As String, Ccy As String, _
              SwaptionVolParameters

1         On Error GoTo ErrHandler
2         Ccy = Right(sh.Name, 3)
3         Set VolRange = sexpandRightDown(RangeFromSheet(sh, "VolInit").Cells(0, 0))
4         Set VolRange = VolRange.Offset(1, 1).Resize(VolRange.Rows.Count - 1, VolRange.Columns.Count - 1)
5         SwaptionVolParameters = RangeFromSheet(sh, "SwaptionVolParameters")
6         QuoteType = sVLookup("QuoteType", SwaptionVolParameters)
7         Contributor = sVLookup("Contributor", SwaptionVolParameters)
8         FeedSwaptionVols VolRange, Ccy, QuoteType, Contributor, Live, AsOfDate
9         Exit Sub
ErrHandler:
10        Throw "#FeedSwaptionVols (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedSwapsOnSheet
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Wrapper to FeedSwapRates. Handles the swap rates on one sheet - i.e. pastes the necessary calls to BDH\BDP to the Hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedSwapsOnSheet(sh As Worksheet, Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler
          Dim Ccy As String
          Dim FixFreqs
          Dim FloatFreqs
          Dim Headers As Variant
          Dim MatchRes As Variant
          Dim RatesRange As Range
          Dim SwapsRange As Range
          Dim Tenors As Variant

2         Ccy = Right(sh.Name, 3)
3         Set SwapsRange = sExpandDown(RangeFromSheet(sh, "SwapRatesInit"))
4         If SwapsRange.Columns.Count <= 6 Then
5             Set SwapsRange = SwapsRange.Resize(, 7)
6         End If

7         Headers = sArrayTranspose(SwapsRange.Rows(0).Value)

8         MatchRes = sMatch("Rate", Headers)
9         If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'Rate' above range of swaps data (" + SwapsRange.Parent.Name + "!" + SwapsRange.Address + ")"
10        Set RatesRange = SwapsRange.Columns(MatchRes)

11        MatchRes = sMatch("Tenor", Headers)
12        If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'Tenor' above range of swaps data (" + SwapsRange.Parent.Name + "!" + SwapsRange.Address + ")"
13        Tenors = SwapsRange.Columns(MatchRes).Value

14        MatchRes = sMatch("FixFreq", Headers)
15        If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'FixFreq' above range of swaps data (" + SwapsRange.Parent.Name + "!" + SwapsRange.Address + ")"
16        FixFreqs = SwapsRange.Columns(MatchRes).Value

17        MatchRes = sMatch("FloatFreq", Headers)
18        If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'FloatFreq' above range of swaps data (" + SwapsRange.Parent.Name + "!" + SwapsRange.Address + ")"
19        FloatFreqs = SwapsRange.Columns(MatchRes).Value

20        FeedSwapRates RatesRange, Tenors, FixFreqs, FloatFreqs, Live, AsOfDate, "SwapRates"

21        Exit Sub
ErrHandler:
22        Throw "#FeedSwapsOnSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedBasisSwapsOnSheet
' Author    : Hermione Glyn
' Date      : 14-Jun-2016
' Purpose   : Wrapper to FeedSwapRates. Handles the basis swap rates on one sheet - i.e. pastes the necessary calls to BDH\BDP to the Hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedBasisSwapsOnSheet(sh As Worksheet, Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler
          Dim BasisSwapsRange As Range
          Dim Ccy As String
          Dim Headers As Variant
          Dim MatchRes As Variant
          Dim RatesRange As Range
          Dim Tenors As Variant

2         Ccy = Right(sh.Name, 3)
3         Set BasisSwapsRange = sExpandDown(RangeFromSheet(sh, "XccyBasisSpreadsInit"))

4         Headers = sArrayTranspose(BasisSwapsRange.Rows(0).Value)

5         MatchRes = sMatch("Rate", Headers)
6         If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'Rate' above range of basis swaps data (" + BasisSwapsRange.Parent.Name + "!" + BasisSwapsRange.Address + ")"
7         Set RatesRange = BasisSwapsRange.Columns(MatchRes)

8         MatchRes = sMatch("Tenor", Headers)
9         If Not IsNumber(MatchRes) Then Throw "Cannot find column headed 'Tenor' above range of basis swaps data (" + BasisSwapsRange.Parent.Name + "!" + BasisSwapsRange.Address + ")"
10        Tenors = BasisSwapsRange.Columns(MatchRes).Value

11        FeedSwapRates RatesRange, Tenors, "Ignored", "Ignored", Live, AsOfDate, "BasisSwapRates"

12        Exit Sub
ErrHandler:
13        Throw "#FeedBasisSwapsOnSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Avg
' Author    : Philip Swannell
' Date      : 10-Jun-2016
' Purpose   : The average of two numbers, with feed-thru of non numbers, use on hidden
'             sheet for converting BDH calls for mid and offer to a mid price
' -----------------------------------------------------------------------------------------------------------------------
Function Avg(a, b)
1         If Not IsNumber(a) Then
2             Avg = "error: " + NonStringToString(a)
3         ElseIf Not IsNumber(b) Then
4             Avg = "error: " + NonStringToString(b)
5         Else
6             Avg = (a + b) / 2
7         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Divide
' Author    : Philip Swannell
' Date      : 11-Jul-2016
' Purpose   : a/b with pass-through of non-numeric inputs to avoid masking errors
' -----------------------------------------------------------------------------------------------------------------------
Function Divide(a, b)
1         On Error GoTo ErrHandler
2         If Not IsNumber(a) Then
3             Divide = "error: " + NonStringToString(a)
4         ElseIf Not IsNumber(b) Then
5             Divide = "error: " + NonStringToString(b)
6         Else
7             Divide = (a / b)
8         End If
9         Exit Function
ErrHandler:
10        Divide = "#Divide (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DateStringToNumber
' Author    : Philip Swannell
' Date      : 11-Jul-2016
' Purpose   : Function BDP returns dates as strings. This function converts to numbers
'             but leaves unchanged any invalid inputs (so that method CalcIsFinished still works)
' -----------------------------------------------------------------------------------------------------------------------
Function DateStringToNumber(a)
          Dim b As Long
1         On Error Resume Next
2         b = CLng(CDate(a))
3         On Error GoTo 0
4         If b = 0 Then
5             DateStringToNumber = a
6         Else
7             DateStringToNumber = b
8         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ValidateBloombergCurveID
' Author    : Philip Swannell
' Date      : 29-Jun-2016
' Purpose   : We need to ensure that the CurveIDs refer at least to curves of the correct currency...
' -----------------------------------------------------------------------------------------------------------------------
Function ValidateBloombergCurveID(CurveID As Long, Ccy As String, Description As String)
          Dim AllCurveIDs As Variant
          Dim BloombergCcy As String
          Dim MatchID As Variant
          Dim Mnemonics As Variant
1         On Error GoTo ErrHandler
2         If CurveID <> 0 Then
3             Mnemonics = sExpandDown(RangeFromSheet(shStaticData, "CurveMnemonic").Offset(1)).Value
4             AllCurveIDs = sExpandDown(RangeFromSheet(shStaticData, "BloombergCurveID").Offset(1)).Value
5             MatchID = sMatch(CurveID, AllCurveIDs)
6             If Not IsNumber(MatchID) Then Throw "Unrecognised " + Description + " for currency " + Ccy
7             BloombergCcy = Left(Mnemonics(MatchID, 1), 3)
8             If BloombergCcy <> Ccy Then
9                 Throw "The " + Description + "for " + Ccy + " is incorrect. It should be a BloombergCurveID for a " + Ccy + " curve, but instead it's a " + BloombergCcy + " curve"
10            End If
11        End If
12        Exit Function
ErrHandler:
13        Throw "#ValidateBloombergCurveID (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedSwapRates
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : For the swap rates on one sheet pastes the necessary calls to BDP and BDH
'             into the Hidden sheet. Used for both Swap rates and basis swap rates.
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedSwapRates(RatesRange As Range, Tenors As Variant, FixFreqs, FloatFreqs, Live As Boolean, AsOfDate As Long, extraInfo1 As String)
1         On Error GoTo ErrHandler
          Dim Ccy As String
          Dim DateString As String
          Dim Divisor As Double
          Dim Formulas As Variant
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long

2         If Not Live Then DateString = "DATe(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"
3         Ccy = Left(RatesRange.Parent.Name, 3)

4         N = RatesRange.Rows.Count: M = RatesRange.Columns.Count
5         Formulas = sReshape("", N, M)
6         For i = 1 To N
7             For j = 1 To M
8                 If extraInfo1 = "SwapRates" Then
9                     Formulas(i, j) = BloombergFormulaSwapRate(Ccy, CStr(Tenors(i, j)), CStr(FixFreqs(i, j)), CStr(FloatFreqs(i, j)), Live, AsOfDate)
10                    Divisor = 100
11                ElseIf extraInfo1 = "BasisSwapRates" Then
12                    If Ccy = RangeFromSheet(shConfig, "CollateralCcy").Value Then
13                        Divisor = 100
14                    Else
15                        Divisor = 10000
16                    End If
17                    Formulas(i, j) = BloombergFormulaBasisSwapRate(Ccy, CStr(Tenors(i, j)), Live, AsOfDate)
18                End If
19            Next j
20        Next i

21        PasteToHiddenSheet RatesRange, Formulas, Divisor, extraInfo1, Ccy

22        Exit Sub
ErrHandler:
23        Throw "#FeedSwapRates (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedSwaptionVols
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : For the swaption vols on one sheet pastes the necessary calls to BDP and BDH
'             into the Hidden sheet.
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedSwaptionVols(VolRange As Range, Ccy As String, QuoteType As String, _
                     Contributor As String, Live As Boolean, AsOfDate As Long)

1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formulas As Variant
          Dim i As Long
          Dim j As Long

2         Formulas = sReshape("", VolRange.Rows.Count, VolRange.Columns.Count)

3         If Not Live Then DateString = "DATe(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

4         For i = 1 To VolRange.Rows.Count
5             For j = 1 To VolRange.Columns.Count
6                 Formulas(i, j) = BloombergFormulaSwaptionVol(Ccy, CStr(VolRange.Cells(i, 0).Value), CStr(VolRange.Cells(0, j).Value), QuoteType, Contributor, Live, AsOfDate)
7             Next j
8         Next i

9         PasteToHiddenSheet VolRange, Formulas, 10000, "SwaptionVols", Ccy
10        Exit Sub
ErrHandler:
11        Throw "#FeedSwaptionVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FeedAllFxData(Live As Boolean, AsOfDate As Long)
          Dim VolRange As Range
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = ThisWorkbook.Worksheets("FX")
3         Set VolRange = sexpandRightDown(RangeFromSheet(ws, "FxDataTopLeft"))
4         Set VolRange = VolRange.Offset(1, 1).Resize(VolRange.Rows.Count - 1, VolRange.Columns.Count - 1)
5         ClearHiddenSheet
6         FeedFxSpotAndVols VolRange, Live, AsOfDate
7         FeedFromHiddenSheet
8         Exit Sub
ErrHandler:
9         Throw "#FeedAllFxData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedFxSpotAndVols
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedFxSpotAndVols(VolRange As Range, Live As Boolean, AsOfDate As Long)
          Dim DateString As String
          Dim Formulas As Variant
          Dim i As Long
          Dim j As Long
          Dim SpotRange As Range
          Dim VolOnlyRange As Range
1         On Error GoTo ErrHandler
2         If VolRange.Cells(0, 1).Value <> "Spot" Then Throw "Assertion failed. Cell above first column of VolRange must read 'Spot'"

3         Formulas = sReshape("", VolRange.Rows.Count, 1)
4         If Not Live Then DateString = "DATe(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

5         Set SpotRange = VolRange.Columns(1)
6         For i = 1 To SpotRange.Rows.Count
7             Formulas(i, 1) = BloombergFormulaFxSpot(VolRange(i, 0), Live, AsOfDate)
8         Next i

9         PasteToHiddenSheet SpotRange, Formulas, 1, "FxSpots", ""

10        Set VolOnlyRange = VolRange.Offset(, 1).Resize(, VolRange.Columns.Count - 1)
11        Formulas = sReshape("", VolOnlyRange.Rows.Count, VolOnlyRange.Columns.Count)
12        For i = 1 To VolOnlyRange.Rows.Count
13            For j = 1 To VolOnlyRange.Columns.Count
14                Formulas(i, j) = BloombergFormulaFxVol(VolOnlyRange(i, -1), VolOnlyRange(0, j), Live, AsOfDate)
15            Next j
16        Next i

17        PasteToHiddenSheet VolOnlyRange, Formulas, 100, "FxVols", ""

18        Exit Sub
ErrHandler:
19        Throw "#FeedFxSpotAndVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetCellComment
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : Adds a comment to a cell and makes it appear in Calibri 11. Comment must be
'             passed including line feed characters.
' -----------------------------------------------------------------------------------------------------------------------
Function SetCellComment(c As Range, Comment As String)
1         On Error GoTo ErrHandler
2         c.ClearComments
3         c.AddComment
4         c.Comment.Visible = False
5         c.Comment.Text Text:=Comment
6         With c.Comment.Shape.TextFrame
7             .Characters.Font.Name = "Calibri"
8             .Characters.Font.Size = 11
9             .AutoSize = True
10        End With
11        Exit Function
ErrHandler:
12        Throw "#SetCellComment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FirstAvailableCell
' Author    : Philip Swannell
' Date      : 10-Jun-2016
' Purpose   : Returns the next cell to write to in column A of the Hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Private Function FirstAvailableCell() As Range
          Dim o As Object
1         On Error GoTo ErrHandler
2         Set o = shHiddenSheet
3         o.UsedRange    'resets the used Range

4         With shHiddenSheet
5             Set FirstAvailableCell = .Cells(.UsedRange.Row + .UsedRange.Rows.Count + 1, 1)
6         End With
7         Exit Function
ErrHandler:
8         Throw "#FirstAvailableCell (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ClearHiddenSheet
' Author    : Philip Swannell
' Date      : 10-Jun-2016
' Purpose   : Removes all data and names from the hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearHiddenSheet()
          Dim N As Name
          Dim Res As Variant
          Dim SPH As clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shHiddenSheet)
3         For Each N In shHiddenSheet.Names
4             N.Delete
5         Next
6         shHiddenSheet.UsedRange.EntireRow.Delete
7         Res = shHiddenSheet.UsedRange.Rows.Count    'ReSeTS USeD RANGe
8         Exit Sub
ErrHandler:
9         Throw "#ClearHiddenSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteToHiddenSheet
' Author    : Philip Swannell
' Date      : 09-Jun-2016
' Purpose   : Common code to be used by methods that feed SwaptionVols, Rates etc
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteToHiddenSheet(TargetRange As Range, Formulas As Variant, Divisor As Double, extraInfo1 As String, extraInfo2 As String)
1         On Error GoTo ErrHandler
          Dim CopyOferr As String
          Dim FirstCell As Range
          Dim i As Long
          Dim j As Long
          Dim Name1 As String
          Dim Name2 As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TempRange As Range
          Dim TempRange2 As Range
          Const ControlBlockHeight = 7
          Dim WithLeftLabels As Boolean
          Dim WithTopLabels As Boolean

2         Set SPH = CreateSheetProtectionHandler(shHiddenSheet)
3         Set SUH = CreateScreenUpdateHandler()

4         Select Case extraInfo1
          Case "CreditSpreads"
5             WithTopLabels = True
6         Case "FxVols", "SwaptionVols"
7             WithTopLabels = True
8             WithLeftLabels = True
9         Case "FxSpots", "SwapRates", "BasisSwapRates"
10            WithLeftLabels = True
11        End Select

12        Set FirstCell = FirstAvailableCell()
13        FirstCell.Cells(1, 1).Value = "Time"
14        FirstCell.Cells(1, 2).Value = "'" + Format(Now, "dd-mmm-yyyy hh:mm:ss")
15        FirstCell.Cells(2, 1).Value = "Target"
16        FirstCell.Cells(2, 2).Value = TargetRange.Parent.Name + "!" + TargetRange.Address
17        Set TempRange = FirstCell.Cells(ControlBlockHeight + IIf(WithTopLabels, 2, 1), 2).Resize(sNRows(Formulas), sNCols(Formulas))
18        Set TempRange2 = TempRange.Offset(0, TempRange.Columns.Count + 1)
19        FirstCell.Cells(3, 1).Value = "Source"
20        FirstCell.Cells(3, 2).Value = TempRange.Address
21        FirstCell.Cells(4, 1).Value = "extraInfo1"
22        FirstCell.Cells(4, 2).Value = extraInfo1
23        FirstCell.Cells(5, 1).Value = "extraInfo2"
24        FirstCell.Cells(5, 2).Value = extraInfo2
25        FirstCell.Cells(6, 1).Value = "FormulasAsText"
26        FirstCell.Cells(6, 2).Value = TempRange2.Address
27        FirstCell.Cells(7, 1).Value = "Pasted"
28        FirstCell.Cells(7, 2).Value = False

29        If WithTopLabels Then
30            TempRange.Rows(0).Value = TargetRange.Rows(0).Value
31        End If
32        If WithLeftLabels Then
33            If extraInfo1 = "FxVols" Then
34                TempRange.Columns(0).Value = TargetRange.Columns(-1).Value
35            Else
36                TempRange.Columns(0).Value = TargetRange.Columns(0).Value
37            End If
38        End If

39        Name1 = "ControlBlock" & (shHiddenSheet.Names.Count / 2) + 1
40        Name2 = "Formulas" & (shHiddenSheet.Names.Count / 2) + 1
41        shHiddenSheet.Names.Add Name1, FirstCell.Resize(ControlBlockHeight, 2)
42        shHiddenSheet.Names.Add Name2, TempRange

43        For i = 1 To TempRange.Rows.Count
44            For j = 1 To TempRange.Columns.Count
45                If Divisor = 1 Then
46                    TempRange.Cells(i, j).Formula = "=" & Formulas(i, j)
47                Else
48                    TempRange.Cells(i, j).Formula = "=Divide(" & Formulas(i, j) & "," & Divisor & ")"
49                End If
50                If i = 1 And j = 1 Then
51                    With TempRange.Cells(i, j)
52                        If IsError(.Value) Then
53                            If .Text = "#NAMe?" Then
54                                .ClearContents
                                  'We assume that the formula evaluating to #NAMe? is because Bloomberg is not installed
                                  'To do -write stand-alone function to test for Bloomberg being both installed and working...
55                                Throw "Feeding rates from Bloomberg is not possible because the Bloomberg addin not installed.", True
56                            End If
57                        End If
58                    End With
59                End If
60            Next j
61        Next i
          'Paste the formulas as text to a range to the right - necessary since method PasteFinishedBlocksFromHiddenSheet _
           replaces formulas with their values once they "resolve" but for debugging we want to have easy access to what those formulas were...
62        TempRange2.FormulaR1C1 = "=FORMULATeXT(RC[" & CStr(TempRange.Column - TempRange2.Column) & "])"
63        TempRange2.Value = sArrayexcelString(TempRange2.Value)

64        Exit Sub
ErrHandler:
65        CopyOferr = "#PasteToHiddenSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
66        Throw CopyOferr
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteFromHiddenSheet
' Author    : Philip Swannell
' Date      : 10-Jun-2016
' Purpose   : Pastes one "block" of the results from calls to BDH or BDP from the hidden sheet back to
'             the correct place in the workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteFromHiddenSheet(ControlBlock As Range, FormulasRange As Range)
          Dim ColorBad
          Dim ColorInterp
          Dim DataToPaste As Variant
          Dim DoInterp As Boolean
          Dim extraInfo1
          Dim extraInfo2
          Dim InterpolatedValues
          Dim SPH As clsSheetProtectionHandler
          Dim TargetAddress As String
          Dim TargetRange As Range

1         On Error GoTo ErrHandler
2         ColorBad = RGB(255, 221, 221)
3         ColorInterp = 11854022    'greenish

4         TargetAddress = sVLookup("Target", ControlBlock.Value)
5         If sIserrorString(TargetAddress) Then Throw "Unexpected error: Cannot find label 'Target' in ControlBlock"

6         extraInfo1 = sVLookup("extraInfo1", ControlBlock)
7         extraInfo2 = sVLookup("extraInfo2", ControlBlock)

8         StatusBarWrap "Pasting Bloomberg results: " + CStr(extraInfo1) + " " + CStr(extraInfo2)

9         Set TargetRange = ThisWorkbook.Worksheets(sStringBetweenStrings(TargetAddress, , "!")).Range(sStringBetweenStrings(TargetAddress, "!"))

10        Set SPH = CreateSheetProtectionHandler(TargetRange.Parent)

11        If FormulasRange.Rows.Count <> TargetRange.Rows.Count Then Throw "Assertion failed: FormulasRange and TargetRange have different number of rows"

12        If FormulasRange.Columns.Count <> TargetRange.Columns.Count Then Throw "Assertion failed: FormulasRange and TargetRange have different number of columns"

13        Select Case extraInfo1
          Case "FxVols", "CreditSpreads"
14            DoInterp = True
15            InterpolatedValues = InterpolateFXVols(FormulasRange.Value, FormulasRange.Rows(0).Value)
16        Case "SwapRates", "BasisSwapRates"
17            DoInterp = True
18            InterpolatedValues = InterpolateSwaps(FormulasRange.Value, FormulasRange.Columns(0).Value)
19        Case "SwaptionVols"
20            DoInterp = True
21            InterpolatedValues = InterpolateSwaptions(FormulasRange.Value, sArrayTranspose(FormulasRange.Rows(0).Value), FormulasRange.Columns(0).Value)
22        End Select

          Dim Formulas As Variant
23        Formulas = FormulasRange.Formula
24        DataToPaste = FormulasRange.Value

25        PasteAndFormat TargetRange, DataToPaste, InterpolatedValues, DoInterp, False, Formulas, , CStr(extraInfo1), CStr(extraInfo2)

26        Exit Sub
ErrHandler:
27        Throw "#PasteFromHiddenSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TenureStringsToTime
' Author    : Philip Swannell
' Date      : 07-Jul-2016
' Purpose   : Simple conversion of an array of tenure strings to time in years
' -----------------------------------------------------------------------------------------------------------------------
Function TenureStringsToTime(TenureStrings)
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Res As Variant
          Dim Tmp As Double

1         On Error GoTo ErrHandler
2         Res = TenureStrings
3         Force2DArray Res
4         N = sNRows(Res): M = sNCols(Res)
5         For i = 1 To N
6             For j = 1 To M
7                 Tmp = CDbl(Left(TenureStrings(i, j), Len(TenureStrings(i, j)) - 1))
8                 Select Case UCase(Right(TenureStrings(i, j), 1))
                  Case "W"
9                     Res(i, j) = Tmp / 365.25 * 7
10                Case "M"
11                    Res(i, j) = Tmp / 12
12                Case "D"
13                    Res(i, j) = Tmp / 365.25
14                Case "Y"
15                    Res(i, j) = Tmp
16                Case Else
17                    Throw "Unrecognised Label"
18                End Select
19            Next j
20        Next i
21        TenureStringsToTime = Res

22        Exit Function
ErrHandler:
23        Throw "#TenureStringsToTime (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InterpolateFXVols
' Author    : Philip Swannell
' Date      : 07-Jul-2016
' Purpose   : Fill in missing FxVols via interpolation and flat extrapolation on each row of the array
' -----------------------------------------------------------------------------------------------------------------------
Function InterpolateFXVols(VolsWitherrors, Labels)
          Dim ChooseVector
          Dim i As Long
          Dim InterpRes
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NumGood As Long
          Dim Res
          Dim ThisRow
          Dim Times As Variant
          Dim xArrayAscending
          Dim yArray
1         On Error GoTo ErrHandler
2         Times = sArrayTranspose(TenureStringsToTime(Labels))
3         Res = VolsWitherrors
4         Force2DArrayR Res
5         NR = sNRows(Res): NC = sNCols(Res)

6         For i = 1 To sNRows(Res)
7             ThisRow = sSubArray(Res, i, 1, 1)
8             ChooseVector = sArrayIsNumber(ThisRow)
9             NumGood = sArrayCount(ChooseVector)
10            If NumGood < NC And NumGood > 0 Then
11                ChooseVector = sArrayTranspose(ChooseVector)
12                xArrayAscending = sMChoose(Times, ChooseVector)
13                yArray = sMChoose(sArrayTranspose(ThisRow), ChooseVector)
14                If NumGood = 1 Then
15                    For j = 1 To NC
16                        Res(i, j) = yArray(1, 1)
17                    Next j
18                Else
19                    InterpRes = sInterp(xArrayAscending, yArray, Times, "Linear", "FF")
20                    For j = 1 To NC
21                        Res(i, j) = InterpRes(j, 1)
22                    Next j
23                End If
24            End If
25        Next i

26        InterpolateFXVols = Res
27        Exit Function
ErrHandler:
28        InterpolateFXVols = "#InterpolateFXVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InterpolateSwaps
' Author    : Philip Swannell
' Date      : 07-Jul-2016
' Purpose   : Rates to be column array, Labels column array of tenure strings
'             returns Rates with non-numeric entries replaced with interpolated\flat extrapolated values
' -----------------------------------------------------------------------------------------------------------------------
Function InterpolateSwaps(Rates, Labels)
          Dim ChooseVector
          Dim i As Long
          Dim InterpRes
          Dim NR As Long
          Dim NumGood As Long
          Dim Res As Variant
          Dim Times
          Dim xArrayAscending
          Dim yArray

1         On Error GoTo ErrHandler
2         Res = Rates
3         Force2DArray Res
4         NR = sNRows(Res)

5         Times = TenureStringsToTime(Labels)
6         ChooseVector = sArrayIsNumber(Res)
7         NumGood = sArrayCount(ChooseVector)
8         If NumGood > 0 And NumGood < NR Then
9             xArrayAscending = sMChoose(Times, ChooseVector)
10            yArray = sMChoose(Rates, ChooseVector)
11            InterpRes = sInterp(xArrayAscending, yArray, Times, "Linear", "FF")
12            If NumGood = 1 Then
13                For i = 1 To NR
14                    Res(i, 1) = yArray(1, 1)
15                Next i
16            Else
17                Res = InterpRes
18            End If
19        End If

20        InterpolateSwaps = Res
21        Exit Function
ErrHandler:
22        Throw "#InterpolateSwaps (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InterpolateSwaptions
' Author    : Philip Swannell
' Date      : 07-Jul-2016
' Purpose   : Fill in missing swaption vols via interpolation and flat extrapolation on
'             each column of the array followed by interpolation\flat extrapolation on each row
' -----------------------------------------------------------------------------------------------------------------------
Function InterpolateSwaptions(VolsWitherrors, TopLabels, LeftLabels)
          Dim ChooseVector
          Dim exercises
          Dim i As Long
          Dim InterpRes
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NumGood As Long
          Dim Res
          Dim Tenures As Variant
          Dim ThisColumn
          Dim ThisRow
          Dim xArrayAscending
          Dim yArray
1         On Error GoTo ErrHandler
2         Tenures = TenureStringsToTime(TopLabels)
3         exercises = TenureStringsToTime(LeftLabels)

4         Res = VolsWitherrors
5         Force2DArrayR Res

6         NR = sNRows(Res): NC = sNCols(Res)

7         NumGood = sArrayCount(sArrayIsNumber(Res))
8         If NumGood = NR * NC Then
9             InterpolateSwaptions = Res
10            Exit Function
11        End If

12        For i = 1 To NC
13            ThisColumn = sSubArray(Res, 1, i, , 1)
14            ChooseVector = sArrayIsNumber(ThisColumn)
15            NumGood = sArrayCount(ChooseVector)
16            If NumGood < NR And NumGood > 0 Then
17                xArrayAscending = sMChoose(exercises, ChooseVector)
18                yArray = sMChoose(ThisColumn, ChooseVector)
19                If NumGood = 1 Then
20                    For j = 1 To NR
21                        Res(j, i) = yArray(1, 1)
22                    Next j
23                Else
24                    InterpRes = sInterp(xArrayAscending, yArray, exercises, "Linear", "FF")
25                    For j = 1 To NR
26                        Res(j, i) = InterpRes(j, 1)
27                    Next j
28                End If
29            End If
30        Next i

31        For i = 1 To NR
32            ThisRow = sSubArray(Res, i, 1, 1)
33            ChooseVector = sArrayIsNumber(ThisRow)
34            NumGood = sArrayCount(ChooseVector)
35            If NumGood < NC And NumGood > 0 Then
36                ChooseVector = sArrayTranspose(ChooseVector)
37                xArrayAscending = sMChoose(Tenures, ChooseVector)
38                yArray = sMChoose(sArrayTranspose(ThisRow), ChooseVector)
39                If NumGood = 1 Then
40                    For j = 1 To NC
41                        Res(i, j) = yArray(1, 1)
42                    Next j
43                Else
44                    InterpRes = sInterp(xArrayAscending, yArray, Tenures, "Linear", "FF")
45                    For j = 1 To NC
46                        Res(i, j) = InterpRes(j, 1)
47                    Next j
48                End If
49            End If
50        Next i

51        InterpolateSwaptions = Res
52        Exit Function
ErrHandler:
53        InterpolateSwaptions = "#InterpolateSwaptions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedCreditData
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedCreditData(Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler
2         ClearHiddenSheet
3         FeedCreditSpreads Live, AsOfDate
4         FeedFromHiddenSheet
5         Exit Sub
ErrHandler:
6         Throw "#FeedCreditData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedCreditSpreads
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   : Updated 03-Jan-2017 to feed to new version of the credit sheet created 29 Dec 2016
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedCreditSpreads(Live As Boolean, AsOfDate As Long)
          Dim AllSpreadTenors() As Variant
          Dim AllTickerTenors() As Variant
          Dim DateString As String
          Dim Formulas() As Variant
          Dim i As Long
          Dim j As Long
          Dim MatchTenorIDs() As Variant
          Dim NC As Long
          Dim NR As Long
          Dim TargetRange As Range
          Dim TickerTableRange As Range

1         On Error GoTo ErrHandler

2         If Not Live Then DateString = "DATe(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

3         Set TickerTableRange = sexpandRightDown(shCredit.Range("CDSTickersTopLeft"))
4         With TickerTableRange
5             Set TickerTableRange = .Offset(1, 0).Resize(.Rows.Count - 1)
6         End With
7         Set TargetRange = CDSRange(shCredit)

8         AllTickerTenors = sArrayTranspose(TickerTableRange.Rows(0).Value)
9         AllSpreadTenors = sArrayTranspose(TargetRange.Rows(1).Offset(0, 3).Resize(1, TargetRange.Columns.Count - 3).Value)
10        MatchTenorIDs = sMatch(AllSpreadTenors, AllTickerTenors)
11        For i = 1 To sNRows(MatchTenorIDs)
12            If Not IsNumber(MatchTenorIDs(i, 1)) Then Throw "Cannot find column labelled " + AllSpreadTenors(i, 1) + " in the tenors for the Ticker Table"
13        Next i

14        With TargetRange
15            Set TargetRange = .Offset(1, 3).Resize(.Rows.Count - 1, .Columns.Count - 3)
16        End With
17        NR = TargetRange.Rows.Count
18        NC = TargetRange.Columns.Count

19        ReDim Formulas(1 To NR, 1 To NC)
20        For j = 1 To NC
21            For i = 1 To NR
22                If Live Then
23                    Formulas(i, j) = "BDP(""" & TickerTableRange(i, j) & " Curncy""" & ",""PX_LAST"")"
24                Else
25                    Formulas(i, j) = "BDH(""" & TickerTableRange(i, j) & " Curncy""" & ",""PX_LAST""," & DateString & "," & DateString & ")"
26                End If
27            Next i
28        Next j

29        PasteToHiddenSheet TargetRange, Formulas, 10000, "CreditSpreads", ""
30        Exit Sub
ErrHandler:
31        Throw "#FeedCreditSpreads (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedInflationSwaps
' Author    : Hermione Glyn
' Date      : 02-May-2017
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedInflationSwaps(Live As Boolean, AsOfDate As Long)
          Dim DateString As String
          Dim Divisor As Double
          Dim Formulas As Variant
          Dim i As Long
          Dim NR As Long
          Dim RatesRange As Range
          Dim Tenors As Variant
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             If IsInflationSheet(ws) Then
4                 If Not Live Then DateString = "DATe(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"
5                 Set RatesRange = sExpandDown(RangeFromSheet(ws, "ZCSwapsInit")).Columns(2)
6                 Tenors = sExpandDown(RangeFromSheet(ws, "ZCSwapsInit")).Columns(1).Value2
7                 NR = sNRows(sExpandDown(RangeFromSheet(ws, "ZCSwapsInit")))
8                 Formulas = sReshape("", NR, 1)
9                 For i = 1 To NR
10                    Formulas(i, 1) = BloombergFormulaZCInflationSwap(ws.Name, CStr(Tenors(i, 1)), Live, AsOfDate)
11                    Divisor = 100
12                Next i
13                PasteToHiddenSheet RatesRange, Formulas, Divisor, "", ""
14            End If
15        Next

16        Exit Sub
ErrHandler:
17        Throw "#FeedInflationSwaps (line " & CStr(Erl) + "): " & Err.Description & "!"
18    End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedRatesFromTextFile
' Author    : Hermione Glyn
' Date      : 30-Sep-2016
' Purpose   : Alternative option for market data feed, allowing the use of a text file
'             instead of Bloomberg. Works for FX sheet and swaps, basis swaps and swaption
'             vols. Uses ParseBBGFile to create an Array of
'             Bloomberg tickers and values.
' PGS 21 March 2022, made into a function so that this method can be called from Cayley
'            via Application.Run
'
' CHANGING THe SIGNATURe OF THIS FUNCTION? ReMeMBeR TO CHeCK THe CALL FROM CAYLeY2022.XLSM!!
'
' -----------------------------------------------------------------------------------------------------------------------
Function FeedRatesFromTextFile(FileName As String, Optional WhatToRefresh As String = "All", _
          Optional ThrowOnerror As Boolean = True)
          
          Dim AnchorDate As Long
          Dim LookUpTable As Variant
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim ws As Worksheet
          Dim ErrString As String
          Dim CCys As String

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler()
3         LookUpTable = ParseBBGFile(FileName, AnchorDate)
4         Set SPH = CreateSheetProtectionHandler(shConfig)
5         RangeFromSheet(shConfig, "AnchorDate").Value2 = AnchorDate

6         If Left(WhatToRefresh, 6) = "Cayley" Then
7             CCys = Mid(WhatToRefresh, 7)

8             For Each ws In ThisWorkbook.Worksheets
9                 Set SPH = CreateSheetProtectionHandler(ws)

10                If IsCurrencySheet(ws) Then
11                    If InStr(CCys, ws.Name) > 0 Then
12                        StatusBarWrap "Updating market data on worksheet " + ws.Name + " from file."    '+ FileName
13                        FeedSheetFromTextFile ws, LookUpTable, "SwapRatesInit"
14                        FeedSheetFromTextFile ws, LookUpTable, "XccyBasisSpreadsInit"
15                        FeedSheetFromTextFile ws, LookUpTable, "VolInit"
16                        FormatCurrencySheet ws, False, Empty
17                    End If
18                ElseIf UCase(ws.Name) = "FX" Then
19                    StatusBarWrap "Updating market data on worksheet " + ws.Name + " from file."    '+ FileName
20                    FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Spot"
21                    FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Vols"
22                    FormatFxVolSheet False
23                End If
24            Next ws

25        Else
26            Select Case WhatToRefresh
                  Case "All"
27                    For Each ws In ThisWorkbook.Worksheets
28                        Set SPH = CreateSheetProtectionHandler(ws)
29                        StatusBarWrap "Updating market data on " + ws.Name + " from file."    '+ FileName
30                        If IsCurrencySheet(ws) Then
31                            FeedSheetFromTextFile ws, LookUpTable, "SwapRatesInit"
32                            FeedSheetFromTextFile ws, LookUpTable, "XccyBasisSpreadsInit"
33                            FeedSheetFromTextFile ws, LookUpTable, "VolInit"
34                            FormatCurrencySheet ws, False, Empty
35                        ElseIf UCase(ws.Name) = "FX" Then
36                            FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Spot"
37                            FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Vols"
38                            FormatFxVolSheet False
39                        ElseIf ws.Name = "Credit" Then
40                            FeedSheetFromTextFile ws, LookUpTable, "CDSTopLeft", ""
41                        ElseIf IsInflationSheet(ws) Then
                              'FeedSheetFromTextFile does not handle inflation (PGS, 24-01-2022)
                              '    FeedSheetFromTextFile ws, LookUpTable, "ZCSwapsInit", ""
42                        End If
43                    Next
44                Case "All Ccys"
45                    For Each ws In ThisWorkbook.Worksheets
46                        Set SPH = CreateSheetProtectionHandler(ws)
47                        If IsCurrencySheet(ws) Then
48                            StatusBarWrap "Updating market data on " + ws.Name + " from file."    '+ FileName
49                            FeedSheetFromTextFile ws, LookUpTable, "SwapRatesInit"
50                            FeedSheetFromTextFile ws, LookUpTable, "XccyBasisSpreadsInit"
51                            FeedSheetFromTextFile ws, LookUpTable, "VolInit"
52                            FormatCurrencySheet ws, False, Empty
53                        End If
54                    Next
55                Case "One Ccy"
56                    Set ws = ThisWorkbook.ActiveSheet
57                    Set SPH = CreateSheetProtectionHandler(ws)
58                    FeedSheetFromTextFile ws, LookUpTable, "SwapRatesInit"
59                    FeedSheetFromTextFile ws, LookUpTable, "XccyBasisSpreadsInit"
60                    FeedSheetFromTextFile ws, LookUpTable, "VolInit"
61                    FormatCurrencySheet ActiveSheet, False, Empty
62                Case "Fx Only"
63                    Set ws = ThisWorkbook.ActiveSheet
64                    Set SPH = CreateSheetProtectionHandler(ws)
65                    FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Spot"
66                    FeedSheetFromTextFile ws, LookUpTable, "FxDataTopLeft", "Vols"
67                    FormatFxVolSheet False
68                Case "Credit Only"
69                    Set ws = ThisWorkbook.ActiveSheet
70                    Set SPH = CreateSheetProtectionHandler(ws)
71                    FeedSheetFromTextFile ws, LookUpTable, "CDSTopLeft", ""
72                Case "Inflation Only"
73                    Set ws = ThisWorkbook.ActiveSheet
74                    Set SPH = CreateSheetProtectionHandler(ws)
75                    FeedSheetFromTextFile ws, LookUpTable, "ZCSwapsInit", ""
76                Case Else
77                    Throw "Did not recognise what to refresh: Must be All, All Ccys, One Ccy, Fx Only, Inflation Only or Credit Only"
78            End Select
79        End If
80        StatusBarWrap False
81        Exit Function

ErrHandler:
82        ErrString = "#FeedRatesFromTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
83        If ThrowOnerror Then
84            Throw ErrString
85        Else
86            FeedRatesFromTextFile = ErrString
87        End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedSheetFromTextFile
' Author    : Hermione Glyn
' Date      : 30-Sep-2016
' Purpose   : Takes in a range name from a currency or Fx sheet: "SwapRatesInit",
'             "XccyBasisSpreadsInit", "VolInit" or "FxDataTopLeft". In the case of
'             FX sheet, extraInfo is used to discern between "Spot" or "Vols".
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedSheetFromTextFile(ws As Worksheet, LookUpTable As Variant, RateClass As String, Optional extraInfo As String)
          Dim Ccy As String
          Dim Divisor As Long
          Dim DoInterp As Boolean
          Dim exercise() As Variant
          Dim FixFreq() As Variant
          Dim FloatFreq() As Variant
          Dim i As Long
          Dim InterpolatedValues() As Variant
          Dim j As Long
          Dim LookupTableLeftCol
          Dim NR As Long
          Dim PX_LASTCol As Variant
          Dim RateNumber() As Variant
          Dim SwapRates() As Variant
          Dim TargetRange As Range
          Dim Tenor() As Variant
          Dim Tickers() As Variant
1         On Error GoTo ErrHandler

2         LookupTableLeftCol = sSubArray(LookUpTable, 1, 1, , 1)

3         PX_LASTCol = sMatch("PX_LAST", sArrayTranspose(sSubArray(LookUpTable, 1, , 1)))

4         Ccy = Left(ws.Name, 3)
5         DoInterp = True

6         Select Case RateClass
              Case "SwapRatesInit"
7                 Divisor = 100
8                 Set TargetRange = sExpandDown(RangeFromSheet(ws, RateClass)).Columns(2)
9                 NR = sNRows(TargetRange)
10                ReDim Tenor(1 To NR): ReDim FixFreq(1 To NR): ReDim FloatFreq(1 To NR): ReDim Tickers(1 To NR): ReDim SwapRates(1 To NR): ReDim RateNumber(1 To NR)
11                For i = 1 To NR
12                    Tenor(i) = RangeFromSheet(ws, RateClass)(i, 1)
13                    FixFreq(i) = RangeFromSheet(ws, RateClass)(i, 3)
14                    FloatFreq(i) = RangeFromSheet(ws, RateClass)(i, 5)
15                    Tickers(i) = BloombergTickerInterestRateSwap(Ccy, CStr(Tenor(i)), CStr(FixFreq(i)), CStr(FloatFreq(i)))
16                    RateNumber(i) = Application.Match(Tickers(i), LookupTableLeftCol, 0)
17                    If Not IsError(RateNumber(i)) Then
18                        SwapRates(i) = Application.WorksheetFunction.Index(LookUpTable, RateNumber(i), PX_LASTCol)    'Get swap rates
19                    Else
20                        SwapRates(i) = Empty
21                    End If
22                Next i
23                Tickers = sArrayTranspose(Tickers)
24                SwapRates = sArrayTranspose(SwapRates)
25                InterpolatedValues = InterpolateSwaps(SwapRates, sArrayTranspose(Tenor))

26            Case "XccyBasisSpreadsInit"
27                Divisor = 10000
28                Set TargetRange = sExpandDown(RangeFromSheet(ws, RateClass)).Columns(2)
29                NR = sNRows(TargetRange)
30                ReDim Tenor(1 To NR): ReDim Tickers(1 To NR): ReDim SwapRates(1 To NR): ReDim RateNumber(1 To NR)
31                For i = 1 To NR
32                    Tenor(i) = RangeFromSheet(ws, RateClass)(i, 1)
33                    Tickers(i) = BloombergTickerBasisSwap(Ccy, CStr(Tenor(i)))
34                    RateNumber(i) = Application.Match(Tickers(i), LookupTableLeftCol, 0)
35                    If Not IsError(RateNumber(i)) Then
36                        SwapRates(i) = Application.WorksheetFunction.Index(LookUpTable, RateNumber(i), PX_LASTCol)    'Get basis swap rates
37                    Else
38                        SwapRates(i) = Empty
39                    End If
40                Next i
41                Tickers = sArrayTranspose(Tickers)
42                SwapRates = sArrayTranspose(SwapRates)
43                InterpolatedValues = InterpolateSwaps(SwapRates, sArrayTranspose(Tenor))

44            Case "VolInit"
45                Divisor = 10000
46                Set TargetRange = sexpandRightDown(RangeFromSheet(ws, RateClass))
47                NR = sNRows(TargetRange)
48                ReDim Tenor(1 To NR): ReDim Tickers(1 To NR, 1 To sNCols(TargetRange)): ReDim RateNumber(1 To NR, 1 To sNCols(TargetRange)): ReDim SwapRates(1 To sNRows(TargetRange), 1 To sNCols(TargetRange))
49                ReDim exercise(1 To sNCols(TargetRange))
                  Dim Contributor As String
                  Dim QuoteType As String
50                QuoteType = sVLookup("QuoteType", RangeFromSheet(ws, "SwaptionVolParameters"))
51                Contributor = sVLookup("Contributor", RangeFromSheet(ws, "SwaptionVolParameters"))

52                For i = 1 To NR
53                    Tenor(i) = TargetRange.Cells(i, 0).Value
54                    For j = 1 To TargetRange.Columns.Count
55                        exercise(j) = TargetRange.Cells(0, j).Value
56                        Tickers(i, j) = BloombergTickerSwaptionVol(Ccy, CStr(exercise(j)), CStr(Tenor(i)), QuoteType, Contributor)
57                        RateNumber(i, j) = Application.Match(Tickers(i, j), LookupTableLeftCol, 0)
58                        If Not IsError(RateNumber(i, j)) Then
59                            SwapRates(i, j) = LookUpTable(RateNumber(i, j), PX_LASTCol)
60                        Else
61                            SwapRates(i, j) = Empty
62                        End If
63                    Next j
64                Next i
65                Tenor = sArrayTranspose(Tenor)
66                exercise = sArrayTranspose(exercise)
67                InterpolatedValues = InterpolateSwaptions(SwapRates, exercise, Tenor)

68            Case "FxDataTopLeft"
69                Divisor = 1
                  Dim CcyPairs As Range
                  Dim FxVols() As Variant
                  Dim Tenors As Range
70                Set CcyPairs = sExpandDown(RangeFromSheet(shFx, "FxDataTopLeft").Offset(1, 0))
71                Set Tenors = sexpandRight(RangeFromSheet(shFx, "FxDataTopLeft").Offset(0, 2))
72                If extraInfo = "Spot" Then
73                    Set TargetRange = sExpandDown(RangeFromSheet(shFx, "FxDataTopLeft").Offset(1, 1))
74                    NR = sNRows(TargetRange)
75                    ReDim RateNumber(1 To NR): ReDim SwapRates(1 To NR)
76                    Tickers = sReshape("", sNRows(CcyPairs), 1)
77                    For i = 1 To sNRows(CcyPairs)
78                        Tickers(i, 1) = BloombergTickerFxSpot(CcyPairs(i, 1))
79                        RateNumber(i) = Application.Match(Tickers(i, 1), LookupTableLeftCol, 0)
80                        If Not IsError(RateNumber(i)) Then
81                            SwapRates(i) = Application.WorksheetFunction.Index(LookUpTable, RateNumber(i), PX_LASTCol)   'Get Fx rates
82                        Else
83                            SwapRates(i) = Empty
84                        End If
85                    Next i
86                    SwapRates = sArrayTranspose(SwapRates)
87                    DoInterp = False

88                ElseIf extraInfo = "Vols" Then
89                    Divisor = 100
90                    Set TargetRange = sexpandRightDown(RangeFromSheet(shFx, "FxDataTopLeft").Offset(1, 2))
91                    NR = sNRows(TargetRange)
92                    ReDim RateNumber(1 To NR, 1 To sNCols(TargetRange)): ReDim SwapRates(1 To sNRows(TargetRange), 1 To sNCols(TargetRange))
93                    Tickers = sReshape("", sNRows(CcyPairs), sNCols(Tenors))
94                    FxVols = sReshape("", sNRows(CcyPairs), sNCols(Tenors))
95                    For i = 1 To sNRows(CcyPairs)
96                        For j = 1 To sNCols(Tenors)
97                            Tickers(i, j) = BloombergTickerFxVol(CcyPairs(i, 1), Tenors(1, j))
98                            FxVols(i, j) = sVLookup(Tickers(i, j), LookUpTable, "PX_LAST", 1)
99                        Next j
100                   Next i
101                   SwapRates = FxVols
102                   DoInterp = True
103                   InterpolatedValues = InterpolateFXVols(FxVols, Tenors)
104               Else
105                   Throw "Range to be refreshed on FX sheet must be Spot or Vols"
106               End If

107           Case "CDSTopLeft"
                  Dim AllSpreadTenors() As Variant
                  Dim AllTickerTenors() As Variant
                  Dim CDSData() As Variant
                  Dim MatchTenorIDs() As Variant
                  Dim NC As Long
                  Dim TickerTableRange As Range
108               Set TickerTableRange = sexpandRightDown(RangeFromSheet(shCredit, "CDSTickersTopLeft").Offset(1, 0))
109               Set TargetRange = CDSRange(shCredit).Offset(1, 3).Resize(TickerTableRange.Rows.Count, TickerTableRange.Columns.Count)
110               NR = sNRows(TargetRange)
111               NC = sNCols(TargetRange)
                  'Check that the tenors are the same
112               AllTickerTenors = sArrayTranspose(TickerTableRange.Rows(0).Value)
113               AllSpreadTenors = sArrayTranspose(TargetRange.Rows(0).Value)
114               MatchTenorIDs = sMatch(AllSpreadTenors, AllTickerTenors)
115               For i = 1 To sNRows(MatchTenorIDs)
116                   If Not IsNumber(MatchTenorIDs(i, 1)) Then Throw "Cannot find column labelled " + AllSpreadTenors(i, 1) + " in the tenors for the Ticker Table"
117               Next i
118               Divisor = 10000
119               ReDim RateNumber(1 To NR, 1 To NC): ReDim SwapRates(1 To NR, 1 To NC)
120               Tickers = sReshape("", NR, NC)
121               CDSData = sReshape("", NR, NC)
122               For i = 1 To NR
123                   For j = 1 To NC
124                       Tickers(i, j) = TickerTableRange(i, j).Value & " Curncy"
125                       CDSData(i, j) = sVLookup(Tickers(i, j), LookUpTable, "PX_LAST", 1)
126                   Next j
127               Next i
128               SwapRates = CDSData
129               DoInterp = True
130               InterpolatedValues = InterpolateFXVols(CDSData, TargetRange.Rows(0).Value)
131           Case Else
132               Throw "Unrecognised RateClass: '" & CStr(RateClass) & "'"
133       End Select

134       PasteAndFormat TargetRange, SwapRates, InterpolatedValues, DoInterp, True, Tickers, Divisor

135       Exit Sub
ErrHandler:
136       Throw "#FeedSheetFromTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteAndFormat
' Author    : Philip Swannell, Hermione Glyn
' Date      : 04-Oct-2016
' Purpose   : Stolen from PasteFromHiddenSheet to share with FeedSheetFromTextFile.
'             Deals with colours, inputs interpolated values and adds comments.
'             extraCommentInfo can be Tickers if FromTextFile is True, should be
'             Formulas if using Bloomberg.
' -----------------------------------------------------------------------------------------------------------------------
Sub PasteAndFormat(TargetRange As Range, DataToPaste As Variant, InterpolatedValues As Variant, DoInterp As Boolean, _
          FromTextFile As Boolean, extraCommentInfo As Variant, Optional Divisor As Long = 1, Optional extraInfo1 As String, Optional extraInfo2 As String)
          Dim c As Range
          Dim ColourBad
          Dim ColourInterp
          Dim extraCommentString As String
          Dim i As Long
          Dim j As Long
          Dim Failures, NumFailures As Long, uniqueFailures, FailureIndicators, uniqueFailuresLeftCol, MatchRes, _
              Comment As String, NumSuccesses As Long, ShowCommentWhenSuccess As Boolean, ThisInterp As Variant

1         On Error GoTo ErrHandler

2         If TargetRange.Parent.Visible <> xlSheetVisible Then
3             TargetRange.Parent.Visible = xlSheetVisible
4         End If

5         ColourBad = RGB(255, 221, 221)
6         ColourInterp = 11854022    'greenish

          'SPeCIAL HANDLING FOR INFLATION INDeX 1 OF 2
7         If extraInfo1 = "InflationIndex" Then    'Avoid pasting non-numbers cauised by data not yet been published
              Dim NumToPaste As Long
8             Force2DArray DataToPaste
9             For i = 1 To sNRows(DataToPaste)
10                If IsNumber(DataToPaste(i, 1)) Then
11                    NumToPaste = NumToPaste + 1
12                Else
13                    Exit For
14                End If
15            Next i
16            If NumToPaste = 0 Then
17                Exit Sub    'nothing to do
18            Else
19                Set TargetRange = TargetRange.Resize(NumToPaste)
20                DataToPaste = sSubArray(DataToPaste, 1, 1, NumToPaste)
21                extraCommentInfo = sSubArray(extraCommentInfo, 1, 1, NumToPaste)
22            End If
23        End If

24        TargetRange.ClearComments
25        TargetRange.Interior.ColorIndex = xlColorIndexAutomatic

26        Failures = DataToPaste

27        NumSuccesses = sArrayCount(sArrayIsNumber(DataToPaste))
28        NumFailures = sNRows(TargetRange) * sNCols(TargetRange) - NumSuccesses

29        If TargetRange.Columns.Count > 1 Then Failures = sReshape(Failures, sNRows(Failures) * sNCols(Failures), 1)

30        If NumFailures > 0 Then
31            Failures = sMChoose(Failures, sArrayNot(sArrayIsNumber(Failures)))
32            uniqueFailures = sCountRepeats(sSortedArray(Failures), "CH")
33            uniqueFailuresLeftCol = sSubArray(uniqueFailures, 1, 1, , 1)
34            FailureIndicators = sReshape(False, sNRows(uniqueFailures), 1)
35        ElseIf NumFailures < 0 Then
36            Throw "NumFailures is less than zero. Check that the target range and input array are of the same size"
37        End If

38        ShowCommentWhenSuccess = True

39        For Each c In TargetRange.Cells
40            i = c.Row - TargetRange.Row + 1
41            j = c.Column - TargetRange.Column + 1

42            If FromTextFile Then
43                extraCommentString = "Ticker: " + extraCommentInfo(i, j)
44            Else
45                extraCommentString = "Formula: " + extraCommentInfo(i, j)
46            End If
47            If DoInterp Then
48                ThisInterp = InterpolatedValues(i, j)
49            End If

50            If IsNumber(DataToPaste(i, j)) Then
51                c.Value = Divide(DataToPaste(i, j), Divisor)
52                If ShowCommentWhenSuccess And Not gApplyRandomAdjustments Then
53                    SetCellComment c, "Feed success:" + vbLf + NonStringToString(DataToPaste(i, j)) + vbLf + extraCommentString + vbLf + "Time:" + vbLf + Format(Now, "d-mmm-yy hh:mm") + _
                          vbLf + "In total there were " + CStr(NumSuccesses) + " successful feeds in this block of cells"
54                    ShowCommentWhenSuccess = False
55                End If
56            Else
57                If IsNumber(ThisInterp) Then
58                    c.Interior.Color = ColourInterp
59                    c.Value = Divide(ThisInterp, Divisor)
60                Else
61                    c.Interior.Color = ColourBad
62                End If
63                MatchRes = sMatch(DataToPaste(i, j), uniqueFailuresLeftCol, True)
64                If Not FailureIndicators(MatchRes, 1) Then
65                    Comment = "Feed failure:" + vbLf + NonStringToString(DataToPaste(i, j)) + vbLf + extraCommentString + vbLf + "Time:" + vbLf + Format(Now, "d-mmm-yy hh:mm")
66                    If uniqueFailures(MatchRes, 2) > 1 Then

67                        Comment = Comment + vbLf + vbLf + _
                              "In total there were " + CStr(uniqueFailures(MatchRes, 2)) + " similar failures in this block of cells."

68                    End If
69                    If DoInterp Then Comment = Comment + vbLf + "Values in green have been interpolated or flat extrapolated" + vbLf + "from values for which a feed was available."
70                    If Not gApplyRandomAdjustments Then
71                        SetCellComment c, Comment
72                    End If
73                End If
74                FailureIndicators(MatchRes, 1) = True
75            End If
76        Next c

          'SPeCIAL HANDLING FOR INFLATION INDeX 2 OF 2
77        If extraInfo1 = "InflationIndex" Then
              Dim M As Long
              Dim Y As Long
78            Y = sStringBetweenStrings(extraInfo2, , ",")
79            M = sStringBetweenStrings(extraInfo2, ",")
80            For i = 1 To NumToPaste
81                TargetRange.Cells(i, -1).Value = Y
82                TargetRange.Cells(i, 0).Value = M
83                Y = IIf(M = 12, Y + 1, Y)
84                M = IIf(M = 12, 1, M + 1)
85            Next i
86            FormatInflationSheet TargetRange.Parent, False, Empty
87        End If

88        Exit Sub

ErrHandler:
89        Throw "#PasteAndFormat (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedFromHiddenSheet
' Author    : Philip
' Date      : 19-Sep-2017
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedFromHiddenSheet()
1         On Error GoTo ErrHandler
2         mCalcCounter = 0
3         CalcHiddenSheet
4         Exit Sub
ErrHandler:
5         Throw "#FeedFromHiddenSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CalcHiddenSheet
' Author    : Philip Swannell
' Date      : 19-Sep-2017
' Purpose   : Calculates the hidden sheet and pastes the results for blocks that have resolved
'             to the appropriate places in the other sheets. If not all blocks have resolved
'             calls itself again in three seconds. Previously tried Doevents to get the BDH
'             calls to resolve but that did not work.
' -----------------------------------------------------------------------------------------------------------------------
Sub CalcHiddenSheet()
1         On Error GoTo ErrHandler
          Dim Res As Boolean
2         mCalcCounter = mCalcCounter + 1
3         StatusBarWrap "Calculate " + CStr(mCalcCounter)
4         shHiddenSheet.Calculate
5         Res = PasteFinishedBlocksFromHiddenSheet()
6         If Res Or mCalcCounter > 60 Then
7             StatusBarWrap False
8             If gApplyRandomAdjustments Then If gDoFx Then AlignFxSpotRates True, RangeFromSheet(shConfig, "Numeraire")
9             Exit Sub
10        Else
11            DoEvents    'Not sure if this helps, but it might possibly...
12            Application.OnTime Now + TimeValue("00:00:03"), ThisWorkbook.Name + "!CalcHiddenSheet"
13        End If
14        Exit Sub
ErrHandler:
15        SomethingWentWrong "#CalcHiddenSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
16        StatusBarWrap False
17        Application.Cursor = xlDefault
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PasteFinishedBlocksFromHiddenSheet
' Author    : Philip Swannell
' Date      : 19-Sep-2017
' Purpose   : Sub of CalcHiddenSheet - pastes blocks to their appropriate place in the workbook
'             sets the "Pasted" flag to TRUe after doing so. Once a block has been pasted back the
'             calls to BDH are replaced by their values, based on a hunch that fewer calls to BDH
'             on a sheet encourage the remaining calls to work.
' -----------------------------------------------------------------------------------------------------------------------
Private Function PasteFinishedBlocksFromHiddenSheet() As Boolean
          Dim c As Range
          Dim N As Name
1         On Error GoTo ErrHandler
          Dim AllDone As Boolean
          Dim ControlBlock As Range
          Dim FlagCell As Range
          Dim IsCalculated As Boolean
          Dim SourceRange As Range
          Dim SourceRangeAddress As String
          Dim SPH As clsSheetProtectionHandler

2         Set SPH = CreateSheetProtectionHandler(shHiddenSheet)
3         AllDone = True
4         For Each N In shHiddenSheet.Names
5             If InStr(N.Name, "ControlBlock") > 0 Then
6                 Set ControlBlock = N.RefersToRange
7                 Set FlagCell = ControlBlock.Cells(sMatch("Pasted", ControlBlock.Columns(1).Value), 2)
8                 If FlagCell.Value = False Then
9                     SourceRangeAddress = sVLookup("Source", ControlBlock.Value)
10                    Set SourceRange = shHiddenSheet.Range(SourceRangeAddress)
11                    DoEvents    'Not sure if this helps, but it might possibly...
12                    SourceRange.Calculate
13                    IsCalculated = True
14                    For Each c In SourceRange.Cells
15                        If InStr(CStr(c.Value2), "Requesting") > 0 Then
16                            IsCalculated = False
17                            Exit For
18                        End If
19                    Next c
20                    If IsCalculated Then
21                        PasteFromHiddenSheet ControlBlock, SourceRange
22                        SourceRange.Value = SourceRange.Value    'Get rid of formulas that have "served their purpose" and reduce the number of calls to BDH\BDP. May help other cells to calculate?
23                        FlagCell.Value = True
24                    Else
25                        AllDone = False
26                    End If
27                End If
28            End If
29        Next N
30        PasteFinishedBlocksFromHiddenSheet = AllDone
31        Exit Function
ErrHandler:
32        Throw "#PasteFinishedBlocksFromHiddenSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FeedInflationIndexes
' Author    : Philip
' Date      : 28-Sep-2017
' Purpose   : Pastes calls to the appropriate Bloomberg function from into the Hidden sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedInflationIndexes(Live As Boolean, ByVal AsOfDate As Long)
          Dim existingData As Range
          Dim FirstMonth As Long
          Dim FirstYear As Long
          Dim Formulas() As String
          Dim i As Long
          Dim LastMonth As Long
          Dim LastYear As Long
          Dim NumMonths
          Dim TargetRange As Range
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         If Live Then AsOfDate = Date

3         For Each ws In ThisWorkbook.Worksheets
4             If IsInflationSheet(ws) Then
5                 Set existingData = sExpandDown(ws.Range("HistoricDataInit"))
6                 For i = existingData.Rows.Count To 1 Step -1
7                     If IsNumber(existingData.Cells(i, 3)) Then
8                         FirstYear = existingData.Cells(i, 1).Value
9                         FirstMonth = existingData.Cells(i, 2).Value
10                        FirstYear = FirstYear + IIf(FirstMonth = 12, 1, 0)
11                        FirstMonth = IIf(FirstMonth = 12, 1, FirstMonth + 1)
12                        Set TargetRange = existingData.Cells(i + 1, 3)
13                        Exit For
14                    End If
15                Next i
16                LastYear = Year(AsOfDate)
17                LastMonth = Month(AsOfDate)
18                NumMonths = 12 * (LastYear - FirstYear) + LastMonth - FirstMonth + 1
19                If NumMonths >= 1 Then
20                    Set TargetRange = TargetRange.Resize(NumMonths)
21                    ReDim Formulas(1 To NumMonths, 1 To 1)
                      Dim M As Long
                      Dim Y As Long
22                    Y = FirstYear: M = FirstMonth
23                    For i = 1 To NumMonths
24                        Formulas(i, 1) = BloombergFormulaInflationIndex(ws.Name, Y, M)
25                        Y = IIf(M = 12, Y + 1, Y)
26                        M = IIf(M = 12, 1, M + 1)
27                    Next i
28                    PasteToHiddenSheet TargetRange, Formulas, 1, "InflationIndex", "'" & FirstYear & "," & FirstMonth
29                End If
30            End If
31        Next

32        Exit Sub
ErrHandler:
33        Throw "#FeedInflationIndexes (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

