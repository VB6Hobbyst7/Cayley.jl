Attribute VB_Name = "modMenu"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowMenu
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Attached to the "Menu..." button on each of the Currency sheets, the Fx sheet and Credit sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowMenu()
          Dim chAddBank As String
          Dim chAddFxVolPair As String
          Dim chAlignFxSpotRates As String
          Dim chClearCommentsActive As String
          Dim chClearCommentsAll As String
          Dim chDeleteBank As String
          Dim chDeleteFxVolPair As String
          Dim chFeedAllCcysFromTextFile As String
          Dim chFeedAllRatesFromTextFile As String
          Dim chFeedCcySheetFromTextFile As String
          Dim chFeedCreditFromTextFile As String
          Dim chFeedFromBloomberg As String
          Dim chFeedFromTextFile As String
          Dim chFeedFXSheetFromTextFile As String
          Dim chImportHistoricalCorr As String
          Dim chOpenCMG As String
          Dim chOpenSCRiPT As String
          Dim Res As String
          Dim TheChoices As Variant
          Dim ThisCCy

1         On Error GoTo ErrHandler
2         Application.EnableCancelKey = xlDisabled
3         ThisCCy = UCase(Right(ActiveSheet.Name, 3))

4         chFeedFromTextFile = "Rates from text &file"
5         chFeedAllRatesFromTextFile = "&All rates, CDS and Fx data"
6         chFeedCcySheetFromTextFile = "&Rates for " + ThisCCy
7         chFeedFXSheetFromTextFile = "Fx rates and vols"
8         chFeedAllCcysFromTextFile = "Rates for all &currencies"
9         chFeedFromBloomberg = "Rates from &Bloomberg..."
10        chAlignFxSpotRates = "Ali&gn Fx Spot Rates..."
11        chClearCommentsActive = "From &this sheet"
12        chClearCommentsAll = "From &all sheets"
13        chAddFxVolPair = "&Add Currency Pair(s)..."
14        chDeleteFxVolPair = "&Delete Currency Pair(s)..."
15        chOpenCMG = "Open Correlation &Matrix Generator"
16        chImportHistoricalCorr = "&Import data from 'Correlation Matrix Generator'"
17        chFeedCreditFromTextFile = "&CDS rates from text file"
18        chAddBank = "&Add a counterparty"
19        chDeleteBank = "&Delete a counterparty"
20        If IsInCollection(Application.Workbooks, sSplitPath(FileFromConfig("SCRiPTWorkbook"), True)) Then
21            chOpenSCRiPT = "Activate SCRiPT &workbook"
22        Else
23            chOpenSCRiPT = "Open SCRiPT &workbook"
24        End If

25        If IsCurrencySheet(ActiveSheet) Then
26            TheChoices = sArrayStack(chFeedFromTextFile, chFeedCcySheetFromTextFile, _
                  "", chFeedAllCcysFromTextFile, _
                  "", chFeedAllRatesFromTextFile, _
                  chFeedFromBloomberg, "", _
                  "--&Clear Bloomberg comments", chClearCommentsActive, _
                  "", chClearCommentsAll)
27            TheChoices = sReshape(TheChoices, sNRows(TheChoices) / 2, 2)
28        ElseIf IsInflationSheet(ActiveSheet) Then
29            TheChoices = sArrayStack(chFeedFromBloomberg, "", _
                  "--&Clear Bloomberg comments", chClearCommentsActive, _
                  "", chClearCommentsAll)

30            TheChoices = sReshape(TheChoices, sNRows(TheChoices) / 2, 2)
31        ElseIf ActiveSheet.Name = shFx.Name Then

32            TheChoices = sArrayStack("Rates from text &file", chFeedFXSheetFromTextFile, _
                  "", chFeedAllCcysFromTextFile, _
                  "", chFeedAllRatesFromTextFile, _
                  chFeedFromBloomberg, "", _
                  chAlignFxSpotRates, "", _
                  "--&Clear Bloomberg comments", chClearCommentsActive, _
                  "", chClearCommentsAll, _
                  "--" & chAddFxVolPair, "", _
                  chDeleteFxVolPair, "")
33            TheChoices = sReshape(TheChoices, sNRows(TheChoices) / 2, 2)
34        ElseIf Left(ActiveSheet.Name, 14) = "HistoricalCorr" Then
35            TheChoices = sArrayStack(chOpenCMG, chImportHistoricalCorr)
36        ElseIf ActiveSheet Is shCredit Then
37            TheChoices = sArrayStack("Feed from text &file", chFeedCreditFromTextFile, _
                  "", chFeedAllRatesFromTextFile, _
                  chFeedFromBloomberg, "", _
                  chAddBank, "", _
                  chDeleteBank, "")
38            TheChoices = sReshape(TheChoices, sNRows(TheChoices) / 2, 2)
39        End If

40        TheChoices = sArrayStack(TheChoices, "--" & chOpenSCRiPT)
          Dim EnableFlags
          Dim MatchID
41        EnableFlags = sReshape(True, sNRows(TheChoices), 1)
42        If Not IsBloombergInstalled Then
43            MatchID = sMatch(chFeedFromBloomberg, sSubArray(TheChoices, 1, 1, , 1))
44            If IsNumber(MatchID) Then
45                EnableFlags(MatchID, 1) = False
46            End If
47        End If

48        Res = ShowCommandBarPopup(TheChoices, , EnableFlags)

49        Application.Cursor = xlWait
50        Select Case Res
              Case "#Cancel!"
51                GoTo earlyExit
52            Case Unembellish(chFeedCcySheetFromTextFile)
53                FeedRatesFromTextFile FileFromConfig("MarketDataFile"), "One Ccy"
54            Case Unembellish(chFeedAllCcysFromTextFile)
55                FeedRatesFromTextFile FileFromConfig("MarketDataFile"), "All Ccys"
56            Case Unembellish(chFeedAllRatesFromTextFile)
57                FeedRatesFromTextFile FileFromConfig("MarketDataFile"), "All"
58            Case Unembellish(chFeedFXSheetFromTextFile)
59                FeedRatesFromTextFile FileFromConfig("MarketDataFile"), "Fx Only"
60            Case Unembellish(chClearCommentsActive)
61                ClearCommentsFromActiveSheet
62            Case Unembellish(chClearCommentsAll)
63                ClearCommentsFromAllSheets
64            Case Unembellish(chAddFxVolPair)
65                AddCurrencyPair
66            Case Unembellish(chDeleteFxVolPair)
67                DeleteRowsFromRange sExpandRightDown(RangeFromSheet(shFx, "FxDataTopLeft")), 1, "Fx Vol", 1, "FxDataTopLeft"
68                SyncHistoricVols
69            Case Unembellish(chImportHistoricalCorr)
70                Application.Cursor = xlDefault
71                ImportHistoricalCorr ActiveSheet
72            Case Unembellish(chOpenSCRiPT)
73                OpenSCRiPTWorkBook False, True
74            Case Unembellish(chOpenCMG)
75                OpenAWorkbook "c:\SolumWorkbooks\Correlation Matrix Generator.xlsm", False, True
76            Case Unembellish(chFeedCreditFromTextFile)
77                FeedRatesFromTextFile FileFromConfig("MarketDataFile"), "Credit Only"
78            Case Unembellish(chFeedFromBloomberg)
79                FeedRatesFromBloomberg
80            Case Unembellish(chAddBank)
81                AddCreditCounterparty
82            Case Unembellish(chDeleteBank)
83                DeleteCreditCounterparty
84            Case Unembellish(chAlignFxSpotRates)
85                AlignFxSpotRates False
86            Case Else
87                Throw "#Unrecognised choice: " & Res
88        End Select
earlyExit:
89        Application.Cursor = xlDefault
90        Exit Sub
ErrHandler:
91        SomethingWentWrong "#ShowMenu (line " & CStr(Erl) + "): " & Err.Description & "!"
92        Application.Cursor = xlDefault
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetCOBDate
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Ask the user what data they want to use when refreshing rates to a date in the past.
' -----------------------------------------------------------------------------------------------------------------------
Function GetCOBDate(ByRef ButtonClicked As String) As Long
          Dim DateAsString As String
          Dim Prompt As String
          Dim Res As Variant
          Static DateAsNumber As Long
          Static LastTime As Double
          Dim ExtraPrompt As String

          Const DateFormat = "dd-mmm-yyyy"
1         On Error GoTo ErrHandler

2         If DateAsNumber = 0 Or (Now() - LastTime) > 0.5 Then
3             Select Case Date Mod 7
                  Case 1
4                     DateAsString = Format(CLng(Date) - 2, DateFormat)
5                 Case 2
6                     DateAsString = Format(CLng(Date) - 3, DateFormat)
7                 Case Else
8                     DateAsString = Format(CLng(Date) - 1, DateFormat)
9             End Select
10        Else
11            DateAsString = Format(DateAsNumber, DateFormat)
12        End If

TryAgain:
13        Application.Cursor = xlDefault
14        Prompt = "Close of business date (" + DateFormat + ")" + vbLf + vbLf + "Note that the AnchorDate on the Config sheet" + vbLf + "is set to be this date." + ExtraPrompt
15        Res = InputBoxPlus(Prompt, "Feed COB data", DateAsString, "< &Back", "&Cancel", , , , , , "&Next >", ButtonClicked)
16        If Res = "" Or Res = False Then
17            GetCOBDate = 0
18            Exit Function
19        End If
20        DateAsNumber = 0
21        On Error Resume Next
22        DateAsNumber = CLng(CDate(Res))
23        On Error GoTo ErrHandler
          'don't allow weekends or dates in the future
24        If DateAsNumber = 0 Then GoTo TryAgain
25        If DateAsNumber > Date Then
26            ExtraPrompt = vbLf + vbLf + "Date entered must be on or before today"
27            GoTo TryAgain
28        End If
29        If DateAsNumber Mod 7 < 2 Then
30            ExtraPrompt = vbLf + vbLf + "Date entered must be a weekday"
31            GoTo TryAgain
32        End If

33        GetCOBDate = DateAsNumber
34        LastTime = Now()

35        Exit Function
ErrHandler:
36        Throw "#GetCOBDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddCurrencyPair
' Author    : Philip Swannell
' Date      : 10-Oct-2016
' Purpose   : Add to the FxVols list on the Fx sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub AddCurrencyPair()
1         On Error GoTo ErrHandler
2         AddPairs "Add Currency Pair", sExpandRightDown(RangeFromSheet(shFx, "FxDataTopLeft")), True
3         SyncHistoricVols
4         Exit Sub
ErrHandler:
5         Throw "#AddCurrencyPair (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddPairs
' Author    : Philip Swannell
' Date      : 15-Jun-2016
' Purpose   : Add a row or rows for a new currency pair to the FX sheet. Called by both AddCurrencyPair and AddFxVol
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddPairs(Title As String, VolRange As Range, HasHeaderRow As Boolean)
1         On Error GoTo ErrHandler
          Dim AllCurrencies As Variant
          Dim ChooseVector As Variant
          Dim ExistingPairs As Variant
          Dim i As Long
          Dim inputBoxRes As String
          Dim newPair As String
          Dim newPair2 As String
          Dim NewPairs As Variant
          Dim Problems As Variant
          Dim Prompt As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim Tmp As Range

2         AllCurrencies = sCurrencies(False, False)

3         If HasHeaderRow Then
4             Set Tmp = VolRange.Offset(1).Resize(VolRange.Rows.Count - 1, 1)
5             ExistingPairs = Tmp.Value
6         Else
7             ExistingPairs = VolRange.Columns(1).Value
8         End If

TryAgain:
9         Prompt = "New currency pair:" + vbLf + "More than one pair? Use a comma-separated list."
10        inputBoxRes = InputBoxPlus(Prompt, Title, inputBoxRes, , , 400)
11        If inputBoxRes = "" Or inputBoxRes = "False" Then GoTo earlyExit

12        NewPairs = sTokeniseString(UCase(inputBoxRes))
13        Force2DArray NewPairs
14        ChooseVector = sReshape(False, sNRows(NewPairs), 1)
15        Problems = sReshape("", sNRows(NewPairs), 1)

16        For i = 1 To sNRows(NewPairs)
17            newPair = UCase(NewPairs(i, 1))
18            newPair2 = Right(newPair, 3) + Left(newPair, 3)
19            If Len(newPair) <> 6 Then
20                ChooseVector(i, 1) = False
21                Problems(i, 1) = "does not have six characters"
22            ElseIf IsNumber(sMatch(newPair, ExistingPairs)) Then
23                ChooseVector(i, 1) = False
24                Problems(i, 1) = "is already listed"
25            ElseIf IsNumber(sMatch(newPair2, ExistingPairs)) Then
26                ChooseVector(i, 1) = False
27                Problems(i, 1) = "is already listed (as " + newPair2 + ")"
28            ElseIf Not IsNumber(sMatch(Left(newPair, 3), AllCurrencies)) Then
29                ChooseVector(i, 1) = False
30                Problems(i, 1) = "is not valid because '" + Left(newPair, 3) + "' is not recognised as a currency"
31            ElseIf Not IsNumber(sMatch(Right(newPair, 3), AllCurrencies)) Then
32                ChooseVector(i, 1) = False
33                Problems(i, 1) = "is not valid because '" + Right(newPair, 3) + "' is not recognised as a currency"
34            ElseIf Left(newPair, 3) = Right(newPair, 3) Then
35                ChooseVector(i, 1) = False
36                Problems(i, 1) = "is not valid because the first and second currencies are the same"
37            Else
38                ChooseVector(i, 1) = True
39                ExistingPairs = sArrayStack(ExistingPairs, newPair)
40                Problems(i, 1) = "is OK to add"
41            End If
42        Next i

43        Problems = sJustifyArrayOfStrings(sArrayRange(NewPairs, Problems), "Calibri", 11, " " & vbTab)

          Dim NumBad As Long
          Dim NumGood As Long
44        NumGood = sArrayCount(ChooseVector)
45        NumBad = sNRows(ChooseVector) - NumGood

46        If NumGood = 0 Then
47            If sNRows(NewPairs) = 1 Then
48                Prompt = Problems(1, 1)
49            Else
50                Prompt = "None of the Fx pairs you typed are valid. Here's why:" + vbLf + sConcatenateStrings(Problems, vbLf)
51            End If
52            MsgBoxPlus Prompt, vbExclamation, Title
53            inputBoxRes = ""
54            GoTo TryAgain
55        End If

56        If NumBad > 0 Then
57            Prompt = "Some of the Fx pairs you typed are not valid. Here's why:" + vbLf + sConcatenateStrings(Problems, vbLf)
58            Prompt = Prompt + vbLf + vbLf + "What do you want to do?" + vbLf + _
                  "Add the valid Fx vols" + vbLf + _
                  "Edit your list of Fx vols" + vbLf + _
                  "Do nothing"
59            Select Case MsgBoxPlus(Prompt, vbQuestion + vbYesNoCancel, Title, "Add valid", "Edit list", "Do nothing", , 400)
                  Case vbYes
                      'continue
60                Case vbNo
61                    inputBoxRes = sConcatenateStrings(sMChoose(NewPairs, ChooseVector))    'just show the good ones
62                    GoTo TryAgain
63                Case Else
64                    Exit Sub
65            End Select
66        End If

67        Set SPH = CreateSheetProtectionHandler(shFx)
68        Set SUH = CreateScreenUpdateHandler()
69        With VolRange
70            .Cells(.Rows.Count + 1, 1).Resize(NumGood, 1).Value = sMChoose(NewPairs, ChooseVector)
71        End With

earlyExit:
72        FormatFxVolSheet False

73        Exit Sub
ErrHandler:
74        SomethingWentWrong "#AddPairs (line " & CStr(Erl) & "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SyncHistoricVols
' Author    : Philip Swannell
' Date      : 30-Mar-2017, updated 31-Jan-2022 to cope with term structure of historic vol.
' Purpose   : We have two ranges of data on the Fx sheet: one for spot and (implied) vol and one for historic vol.
'             This method brings the labels in the left col of the historic vols data into sync with the labels in the implied vols
'             should be called after we add or delete pairs from the implied vols data.
' -----------------------------------------------------------------------------------------------------------------------
Sub SyncHistoricVols()
          Dim ExistingData
          Dim ExistingRange
          Dim FromRow
          Dim i As Long
          Dim j As Long
          Dim LabelsRange
          Dim MatchIDs
          Dim NC As Long
          Dim NewData
          Dim NewDataLeftCol
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set ExistingRange = sExpandRightDown(RangeFromSheet(shFx, "HistoricFxVolsTopLeft"))
3         NC = ExistingRange.Columns.Count
4         ExistingData = ExistingRange.Value

5         Set LabelsRange = sExpandDown(RangeFromSheet(shFx, "FxDataTopLeft"))
6         With LabelsRange
7             NewDataLeftCol = .Offset(1).Resize(.Rows.Count - 1).Value
8         End With

9         If sArraysIdentical(NewDataLeftCol, sSubArray(ExistingData, 1, 1, , 1)) Then
10            Exit Sub
11        End If

12        MatchIDs = sMatch(NewDataLeftCol, sSubArray(ExistingData, 1, 1, , 1))

13        NewData = sReshape("", sNRows(NewDataLeftCol), sNCols(ExistingData))

14        For i = 1 To sNRows(NewData)
15            FromRow = MatchIDs(i, 1)
16            If IsNumber(FromRow) Then
17                For j = 1 To NC
18                    NewData(i, j) = ExistingData(FromRow, j)
19                Next j
20            Else
21                NewData(i, 1) = NewDataLeftCol(i, 1)
22                For j = 2 To NC
23                    NewData(i, j) = Empty
24                Next j
25            End If
26        Next i

27        Set SPH = CreateSheetProtectionHandler(shFx)
28        ExistingRange.Clear
29        ExistingRange.Resize(sNRows(NewData)).Value = NewData
30        FormatFxVolSheet False

31        Exit Sub
ErrHandler:
32        Throw "#SyncHistoricVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function OpenCorrelationMatrixGenerator(HideOnOpening As Boolean, Activate As Boolean) As Workbook
1         On Error GoTo ErrHandler
2         Set OpenCorrelationMatrixGenerator = OpenAWorkbook("c:\SolumWorkbooks\CorrelationMatrixGenerator.xlsm", HideOnOpening, Activate)
3         Exit Function
ErrHandler:
4         Throw "#OpenCorrelationMatrixGenerator (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function OpenSCRiPTWorkBook(HideOnOpening As Boolean, Activate As Boolean) As Workbook
1         On Error GoTo ErrHandler
2         Set OpenSCRiPTWorkBook = OpenAWorkbook(FileFromConfig("SCRiPTWorkbook"), HideOnOpening, Activate)
3         Exit Function
ErrHandler:
4         Throw "#OpenSCRiPTWorkBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenAWorkbook
' Author    : Philip Swannell
' Date      : 3-Dec-2015
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Private Function OpenAWorkbook(FullName As String, HideOnOpening As Boolean, Activate As Boolean) As Workbook
          Dim BookName As String
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         BookName = sSplitPath(FullName)

3         If Not IsInCollection(Application.Workbooks, BookName) Then
4             Set wb = Application.Workbooks.Open(FullName, , False)
5             If HideOnOpening Then
6                 wb.Windows(1).Visible = False
7             End If
8         Else
9             Set wb = Application.Workbooks(BookName)
10        End If

11        If Activate Then
12            With wb.Windows(1)
13                .Visible = True
14                .Activate
15                If .WindowState = xlMinimized Then .WindowState = xlNormal
16            End With
17        Else
18            ThisWorkbook.Activate
19        End If

20        Set OpenAWorkbook = wb

21        Exit Function
ErrHandler:
22        Throw "#OpenAWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DeleteCreditCounterparty
' Author    : Hermione Glyn
' Date      : 16-Jan-2017
' Purpose   : Remove a bank/counterparty from the CDS data and ticker table.
'             DeleteRowsFromRange gives a prompt of all banks shown on LHS column.
' -----------------------------------------------------------------------------------------------------------------------
Sub DeleteCreditCounterparty()
          Dim R As Range
          Dim rowAlias As String
          Dim TopLeftName As String
1         On Error GoTo ErrHandler
2         TopLeftName = "CDSTopLeft"
3         Set R = sExpandRightDown(RangeFromSheet(shCredit, TopLeftName))
4         Set R = R.Offset(-1).Resize(R.Rows.Count + 1)
5         rowAlias = "counterparty"
6         DeleteRowsFromRange R, 2, rowAlias, 1, TopLeftName, "FormatCreditSheet2"
7         Exit Sub
ErrHandler:
8         Throw "#DeleteCreditCounterparty (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddCreditCounterparty
' Author    : Hermione Glyn
' Date      : 16-Jan-2017
' Purpose   : Asks for a counterparty name and uses this to create a new row of CDS data
'             on the Credit sheet.
' -----------------------------------------------------------------------------------------------------------------------
Sub AddCreditCounterparty()
          Dim CDSData As Range
          Dim ChooseVector As Variant
          Dim ExistingCounterparties As Variant
          Dim i As Long
          Dim inputBoxRes As String
          Dim NewCounterparties As Variant
          Dim NewCounterparty As String
          Dim Problems As Variant
          Dim Prompt As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set CDSData = CDSRange(shCredit)
3         ExistingCounterparties = CDSData.Offset(1, 0).Columns(1).Value

TryAgain:
4         Prompt = "New counterparty:" + vbLf + "More than one? Use a comma-separated list."
5         inputBoxRes = InputBoxPlus(Prompt, , inputBoxRes, , , 400)
6         If inputBoxRes = "" Or inputBoxRes = "False" Then GoTo earlyExit

7         NewCounterparties = sTokeniseString(UCase(inputBoxRes))
8         Force2DArray NewCounterparties
9         ChooseVector = sReshape(False, sNRows(NewCounterparties), 1)
10        Problems = sReshape("", sNRows(NewCounterparties), 1)

11        For i = 1 To sNRows(NewCounterparties)
12            NewCounterparty = UCase(NewCounterparties(i, 1))
13            If IsNumber(sMatch(NewCounterparty, ExistingCounterparties)) Then
14                ChooseVector(i, 1) = False
15                Problems(i, 1) = "is already listed"
16            Else
17                ChooseVector(i, 1) = True
18                ExistingCounterparties = sArrayStack(ExistingCounterparties, NewCounterparty)
19                Problems(i, 1) = "is OK to add"
20            End If
21        Next i

22        Problems = sJustifyArrayOfStrings(sArrayRange(NewCounterparties, Problems), "Calibri", 11, " " & vbTab)

          Dim NumBad As Long
          Dim NumGood As Long
23        NumGood = sArrayCount(ChooseVector)
24        NumBad = sNRows(ChooseVector) - NumGood

25        If NumGood = 0 Then
26            If sNRows(NewCounterparties) = 1 Then
27                Prompt = Problems(1, 1)
28            Else
29                Prompt = "None of the counterparties you typed are valid. Here's why:" + vbLf + sConcatenateStrings(Problems, vbLf)
30            End If
31            MsgBoxPlus Prompt, vbExclamation
32            inputBoxRes = ""
33            GoTo TryAgain
34        End If

35        If NumBad > 0 Then
36            Prompt = "Some of the counterparties you typed are not valid. Here's why:" + vbLf + sConcatenateStrings(Problems, vbLf)
37            Prompt = Prompt + vbLf + vbLf + "What do you want to do?" + vbLf + _
                  "Add the valid counterparties" + vbLf + _
                  "Edit your list of counterparties" + vbLf + _
                  "Do nothing"
38            Select Case MsgBoxPlus(Prompt, vbQuestion + vbYesNoCancel, "Add valid", "Edit list", "Do nothing", , 400)
                  Case vbYes
                      'continue
39                Case vbNo
40                    inputBoxRes = sConcatenateStrings(sMChoose(NewCounterparties, ChooseVector))    'just show the good ones
41                    GoTo TryAgain
42                Case Else
43                    Exit Sub
44            End Select
45        End If

46        Set SPH = CreateSheetProtectionHandler(shCredit)
47        Set SUH = CreateScreenUpdateHandler()
48        With CDSData
49            .Cells(.Rows.Count + 1, 1).Resize(NumGood, 1).Value = sMChoose(NewCounterparties, ChooseVector)
50        End With

earlyExit:
51        FormatCreditSheet False
52        Exit Sub
ErrHandler:
53        Throw "#AddCreditCounterparty (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function MakePlural(Widget As String)
1         On Error GoTo ErrHandler
2         If Right(Widget, 1) = "y" Then
3             MakePlural = sArrayLeft(Widget, -1) + "ies"
4         Else
5             MakePlural = Widget + "s"
6         End If
7         Exit Function
ErrHandler:
8         Throw "#MakePlural (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DeleteRowsFromRange
' Author    : Philip Swannell
' Date      : 10-Oct-2016
' Purpose   : General purpose "delete rows from range" method.
'  Assumes: 1) left hand column contains a unique identifier of some kind
'           2) deleting cells with an up shift will not screw up data elsewhere on the sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub DeleteRowsFromRange(R As Range, NumHeaderRows As Long, rowAlias As String, MinAllowed As Long, Optional TopLeftName As String, Optional FormatMethodName As String)
          Dim Choices As Variant
          Dim Chosen As Variant
          Dim EndNum As Long
          Dim i As Long
          Dim MatchIDs As Variant
          Dim OrigTopLeftAddress As String
          Dim Prompt As String
          Dim RangeToDelete As Range
          Dim SPH As clsSheetProtectionHandler
          Dim StartNum As Long
          Dim SUH As clsScreenUpdateHandler
          Dim Title As String

1         On Error GoTo ErrHandler
2         Prompt = "Select " + MakePlural(rowAlias)
3         Title = "Delete " + MakePlural(rowAlias)
4         Choices = R.Columns(1).Offset(NumHeaderRows).Resize(R.Rows.Count - NumHeaderRows).Value
5         StartNum = sNRows(Choices)
TryAgain:
6         Chosen = ShowMultipleChoiceDialog(Choices, , Title, Prompt)
7         If sArraysIdentical(Chosen, "#User Cancel!") Then Exit Sub
8         If IsEmpty(Chosen) Then Exit Sub
9         EndNum = StartNum - sNRows(Chosen)
10        If EndNum < MinAllowed Then
11            MsgBoxPlus "You must leave at least " + CStr(MinAllowed) + " " + rowAlias + " not deleted", , Title
12            GoTo TryAgain
13        End If
14        Prompt = "Delete " + CStr(sNRows(Chosen)) + " " + IIf(sNRows(Chosen) > 1, MakePlural(rowAlias), rowAlias) + "?"    ' + vbLf + vbLf + "Undo with Ctrl Z"

15        If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel + vbDefaultButton2, Title, "Yes, delete", "No, do nothing") <> vbOK Then Exit Sub

16        MatchIDs = sMatch(Chosen, Choices)
17        MatchIDs = sArrayAdd(NumHeaderRows, MatchIDs)
18        Force2DArray MatchIDs

19        Set RangeToDelete = R.Rows(MatchIDs(1, 1))

20        For i = 2 To sNRows(MatchIDs)
21            Set RangeToDelete = Application.Union(RangeToDelete, R.Rows(MatchIDs(i, 1)))
22        Next i

23        Application.GoTo RangeToDelete

          ' BackUpRange R, shUndo, R.Parent.Cells(1, 1)
24        Set SPH = CreateSheetProtectionHandler(R.Parent)
25        Set SUH = CreateScreenUpdateHandler()

26        If TopLeftName <> "" Then
27            OrigTopLeftAddress = R.Parent.Range(TopLeftName).Address
28        End If

29        RangeToDelete.Delete xlShiftUp
30        AddGreyBorders R, True
31        AddGreyBorders R.Resize(, 1), True
32        If NumHeaderRows > 0 Then
33            AddGreyBorders R.Resize(NumHeaderRows), True
34            AddGreyBorders R.Resize(NumHeaderRows, 1), True
35        End If

36        If TopLeftName <> "" Then
37            R.Parent.Names.Add TopLeftName, R.Parent.Range(OrigTopLeftAddress)
38        End If
39        Application.GoTo R.Parent.Cells(1, 1)
40        Application.GoTo R.Cells(1, 1)
41        Set SPH = Nothing    'set to nothing now otherwise its terminate event clear out the undo buffer

42        If FormatMethodName <> "" Then
43            ThrowIfError Application.Run(FormatMethodName)
44        End If

          'Application.OnUndo "Undo " & Title, "'SolumAddin.xlam'!RestoreRange"

45        Exit Sub
ErrHandler:
46        Throw "#DeleteRowsFromRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AlignFxSpotRates
' Author    : Philip Swannell
' Date      : 30-Mar-2017
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub AlignFxSpotRates(Silent As Boolean, Optional BaseCCY As String)
          Dim ChooseVector As Variant
          Dim i As Long
          Dim NewRates
          Dim Pairs As Variant
          Dim Prompt As String
          Dim Rates As Variant
          Dim SpotRange As Range
          Const Threshold = 0.0025
          Dim PromptArray As Variant
          Const Title = "Align Fx Spot Rates"
          Dim Numchanged As Long
          Dim SPH As clsSheetProtectionHandler
          Dim TopText As String

1         On Error GoTo ErrHandler
2         Prompt = "Align spot cross rates with spot rates against a chosen base currency?" + vbLf + vbLf + "You will be warned if any rate would change by more than " + Format(Threshold, "0.00%") + "."

3         If Not Silent Then If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, "&Next >") <> vbOK Then Exit Sub

4         TopText = "Choose base currency." + vbLf + "Cross rates will be inferred from rates against the base currency."

5         If BaseCCY = "" Then
6             BaseCCY = ShowOptionButtonDialog(sArrayStack("USD", "EUR", "GBP"), Title, TopText)
7         End If
8         If BaseCCY = "" Then Exit Sub

9         Set SpotRange = sExpandDown(RangeFromSheet(shFx, "FxDataTopLeft").Offset(1)).Offset(, 1)
10        Pairs = SpotRange.Offset(, -1).Value
11        Rates = SpotRange.Value
12        NewRates = RebaseFX(Pairs, Rates, BaseCCY)

13        ChooseVector = sArrayGreaterThan(sArrayAbs(sArraySubtract(sArrayDivide(NewRates, Rates), 1)), Threshold)

14        If Not Silent Then
15            If sArrayCount(ChooseVector) > 0 Then
                  Dim PercentageChanges
16                PercentageChanges = sMChoose(sArraySubtract(sArrayDivide(NewRates, Rates), 1), ChooseVector)
17                For i = 1 To sNRows(PercentageChanges)
18                    PercentageChanges(i, 1) = Format(PercentageChanges(i, 1), "0.00%")
19                Next i
20                PromptArray = sArrayRange(sMChoose(Pairs, ChooseVector), sMChoose(Rates, ChooseVector), sMChoose(NewRates, ChooseVector), PercentageChanges)
21                PromptArray = sArrayStack(sArrayRange("Ccy Pair", "Old value", "New value", "% increase"), PromptArray)
22                Prompt = "Operation will change the folowing rates by more than " + Format(Threshold, "0.00%") + vbLf + vbLf + _
                      sConcatenateStrings(sJustifyArrayOfStrings(PromptArray, "Calibri", 11, " " & vbTab, False, False), vbLf)
23                Prompt = Prompt + vbLf + vbLf + "Do you want to continue?"
24                If MsgBoxPlus(Prompt, vbYesNo + vbQuestion, , "Yes, continue", "No, do nothing", Title, , 400) <> vbYes Then Exit Sub
25            End If
26        End If
27        Numchanged = sArrayCount(sArrayNot(sArrayEquals(Rates, NewRates)))

28        If Not sArraysIdentical(Rates, NewRates) Then
29            Set SPH = CreateSheetProtectionHandler(shFx)
30            SpotRange.Value = NewRates
31        End If
32        TemporaryMessage "Align Fx Spot Rates: " + CStr(Numchanged) + " rate(s) changed"

33        Exit Sub
ErrHandler:
34        Throw "#AlignFxSpotRates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RebaseFX
' Author    : Philip Swannell
' Date      : 30-Mar-2017
' Purpose   : Sub of AlignFxSpotRates
' -----------------------------------------------------------------------------------------------------------------------
Function RebaseFX(Pairs, Rates, ByVal BaseCCY As String)

          Dim i As Long
          Dim N As Long

          Dim C1 As String
          Dim C1Base As Double
          Dim C1s As Variant
          Dim C2 As String
          Dim C2Base As Double
          Dim C2s As Variant
          Dim CanDo As Boolean
          Dim MatchRes1
          Dim MatchRes2
          Dim MatchRes3
          Dim MatchRes4
          Dim Result As Variant
1         On Error GoTo ErrHandler
2         BaseCCY = UCase(BaseCCY)

3         C1s = sArrayLeft(Pairs, 3)
4         C2s = sArrayRight(Pairs, 3)

5         MatchRes1 = sMatch(sArrayConcatenate(BaseCCY, C1s), Pairs)
6         MatchRes2 = sMatch(sArrayConcatenate(C1s, BaseCCY), Pairs)
7         MatchRes3 = sMatch(sArrayConcatenate(BaseCCY, C2s), Pairs)
8         MatchRes4 = sMatch(sArrayConcatenate(C2s, BaseCCY), Pairs)
9         N = sNRows(Pairs)

10        Result = sReshape("", N, 1)

11        For i = 1 To N
12            CanDo = True
13            C1 = UCase(C1s(i, 1))
14            C2 = UCase(C2s(i, 1))
15            If C1 = BaseCCY Or C2 = BaseCCY Then
16                Result(i, 1) = Rates(i, 1)
                  'Result(i, 2) = Pairs(i, 1)
                  'Result(i, 3) = Rates(i, 1)
17            Else
18                If C1 = BaseCCY Then
19                    C1Base = 1
20                ElseIf IsNumber(MatchRes1(i, 1)) Then
21                    If IsNumber(Rates(MatchRes1(i, 1), 1)) Then
22                        C1Base = 1 / Rates(MatchRes1(i, 1), 1)
                          'Result(i, 2) = BaseCCy & C1
                          'Result(i, 3) = Rates(MatchRes1(i, 1), 1)
23                    Else
24                        CanDo = False
25                    End If
26                ElseIf IsNumber(MatchRes2(i, 1)) Then
27                    If IsNumber(Rates(MatchRes2(i, 1), 1)) Then
28                        C1Base = Rates(MatchRes2(i, 1), 1)
                          'Result(i, 2) = C1 & BaseCCy
                          'Result(i, 3) = Rates(MatchRes2(i, 1), 1)
29                    Else
30                        CanDo = False
31                    End If
32                Else
33                    CanDo = False
34                End If
35                If C2 = BaseCCY Then
36                    C2Base = 1
37                ElseIf IsNumber(MatchRes3(i, 1)) Then
38                    If IsNumber(Rates(MatchRes3(i, 1), 1)) Then
39                        C2Base = 1 / Rates(MatchRes3(i, 1), 1)
                          'Result(i, 4) = BaseCCy & C2
                          'Result(i, 5) = Rates(MatchRes3(i, 1), 1)
40                    Else
41                        CanDo = False
42                    End If
43                ElseIf IsNumber(MatchRes4(i, 1)) Then
44                    If IsNumber(Rates(MatchRes4(i, 1), 1)) Then
45                        C2Base = Rates(MatchRes4(i, 1), 1)
                          'Result(i, 4) = C2 & BaseCCy
                          'Result(i, 5) = Rates(MatchRes4(i, 1), 1)
46                    Else
47                        CanDo = False
48                    End If
49                Else
50                    CanDo = False
51                End If
52                If CanDo Then
53                    Result(i, 1) = C1Base / C2Base
54                Else
55                    Result(i, 1) = Rates(i, 1)
56                End If
57            End If
58        Next i

59        RebaseFX = Result
60        Exit Function
ErrHandler:
61        RebaseFX = "#RebaseFX (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

