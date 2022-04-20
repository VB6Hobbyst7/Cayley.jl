Attribute VB_Name = "modMarketDataLnFx"
'---------------------------------------------------------------------------------------
' Module    : modMarketData
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Code to read the sheets of the market data workbook, format data so that it
'             can be understood by the Julia code, and save that data down to file
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : OpenMarketWorkbook
' Author    : Philip Swannell
' Date      : 3-Dec-2015
' Purpose   : Returns a workbook object representing the market data workbook whose
'             Fullname is given on the Config sheet of this workbook
'---------------------------------------------------------------------------------------
Function OpenMarketWorkbook(Optional HideOnOpening As Boolean, Optional Activate As Boolean) As Workbook
          Dim BookFullName As String
          Dim BookName As String
          Dim sh As Worksheet
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         BookFullName = FileFromConfig("MarketDataWorkbook")
3         BookName = sSplitPath(BookFullName)
4         Set SUH = CreateScreenUpdateHandler()

5         If Not IsInCollection(Application.Workbooks, BookName) Then
6             Set wb = Application.Workbooks.Open(BookFullName, , False)
7             CheckMarketWorkbook wb, gProjectName
8             shConfig.Calculate    'since the Config sheet calls method RangeFromMarketDataBook
9             If HideOnOpening Then
10                wb.Windows(1).Visible = False
11            End If
              'Ensure that the market data workbook that we are opening "points to" this instance of the SCRiPT workbook
12            If IsInCollection(wb.Worksheets, "Config") Then
13                Set sh = wb.Worksheets("Config")
14                If IsInCollection(sh.Names, "SCRiPTWorkbook") Then
15                    If RangeFromSheet(sh, "SCRiPTWorkbook").Value <> ThisWorkbook.FullName Then
16                        Set SPH = CreateSheetProtectionHandler(sh)
17                        RangeFromSheet(sh, "SCRiPTWorkbook").Value = "'" & ThisWorkbook.FullName
18                        AddFileToMRU "SCRiPTWorkbooks", ThisWorkbook.FullName
19                    End If
20                End If
21            End If
22        Else
23            Set wb = Application.Workbooks(BookName)
24        End If

25        If Activate Then
26            With wb.Windows(1)
27                .Visible = True
28                If .WindowState = xlMinimized Then .WindowState = xlNormal
29                .Activate
30            End With
31        Else
32            ThisWorkbook.Windows(1).Visible = True
33            ThisWorkbook.Windows(1).Activate
34        End If

35        Set OpenMarketWorkbook = wb

36        Exit Function
ErrHandler:
37        Throw "#OpenMarketWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CalculateRequiredMarketData
' Author    : Philip Swannell
' Date      : 23-Mar-2016
' Purpose   : From the trades on the Portfolio sheet calculate a list of currencies and
'             inflation indices which need to be in the model. Those two lists can be
'             passed to method SaveDataFromMarketWorkbookToFile. Note that whenever a currency is "required"
'             then, to calibrate the model, we will have to save down swap rates, basis swap rates,
'             vol surface, fx vols vs numeraire. But this is handled in SaveDataFromMarketWorkbookToFile.
'             Also this method does not need to handle the dependency of Inflation indices on a
'             currency - e.g. to calibrate forwards for UKRPI we need GBP to be in the model. That dependency
'             is also handled within SaveDataFromMarketWorkbookToFile.
'---------------------------------------------------------------------------------------
Function CalculateRequiredMarketData(ByRef Ccys As Variant, ByRef IIs As Variant)

          Dim CCysInMWB
          Dim CollateralCcy As String
          Dim NonNumeraireCCysInMWB
          Dim Numeraire As String
          Dim NumInflation As Long
          Dim NumRatesFx As Long
          Dim NumTrades As Long
          Dim TR As Range
          Const PreferenceOrder = "USD,EUR,GBP,JPY"    'Helps alleviate the "add a currency and PFEs of existing trades changes" problem. Numeraire always goes first though

1         On Error GoTo ErrHandler
2         OpenMarketWorkbook True
3         Numeraire = RangeFromMarketDataBook("Config", "Numeraire")
4         CollateralCcy = RangeFromMarketDataBook("Config", "CollateralCcy")
5         Set TR = getTradesRange(NumTrades)

6         If NumTrades = 0 Then
7             Ccys = Numeraire
8             NumRatesFx = 1
9             IIs = CreateMissing()
10            NumInflation = 0
11        Else
              Dim Both
12            Both = sArrayStack(Numeraire, CollateralCcy, TR.Columns(gCN_Ccy1), TR.Columns(gCN_Ccy2))
13            Both = sRemoveDuplicates(Both, True)
14            Both = sMChoose(Both, sIsRegMatch("^((?!^N/A$).)*$", Both))    'get rid of "N/A" that is allowed to appear on the Portfolio sheet
15            IIs = sCompareTwoArrays(Both, SupportedInflationIndices(), "Common")
16            If sNRows(IIs) > 1 Then
                  'Note that we don't worry here whether or not the base currency associated with each inflation index is included in the array Ccys. Method SaveDataFromMarketWorkbookToFile in the market data workbook handles that issue.
17                IIs = sSubArray(IIs, 2)
18                NumInflation = sNRows(IIs)
19            Else
20                IIs = CreateMissing()
21                NumInflation = 0
22            End If
23            Ccys = sCompareTwoArrays(Both, sCurrencies(False, False), "Common,NoHeaders")
24            NumRatesFx = sNRows(Ccys)
25        End If

          'PGS 19-May-2015 R method produceState.xccyHW fails if there is only one currency. Should fix this, but for the time being simply add another currency
26        If sNRows(Ccys) = 1 Then
27            If NumInflation = 0 Then    'because at a deeper level inflation indexes are currencies
28                CCysInMWB = CurrenciesInMarketWorkbook(False)
29                If sNRows(CCysInMWB) = 1 Then Throw "Market data workbook must have sheets for at least two currencies"
                  'exclude the numeraire from the list
30                NonNumeraireCCysInMWB = sMChoose(CCysInMWB, sArrayNot(sArrayEquals(CCysInMWB, Numeraire)))

31                If Numeraire = "USD" Then    'tack on EUR, as long as there is data available for EUR
32                    If IsNumber(sMatch("EUR", CCysInMWB)) Then
33                        Ccys = sArrayStack(Ccys, "EUR")
34                    Else
35                        Ccys = sArrayStack(Ccys, NonNumeraireCCysInMWB(1, 1))
36                    End If
37                Else
38                    If IsNumber(sMatch("USD", CCysInMWB)) Then
39                        Ccys = sArrayStack(Ccys, "USD")
40                    Else
41                        Ccys = sArrayStack(Ccys, NonNumeraireCCysInMWB(1, 1))
42                    End If
43                End If
44            End If
45        End If

          'Reorder
          Dim MatchIDs
46        MatchIDs = sMatch(Ccys, sTokeniseString(Numeraire & "," & PreferenceOrder))
47        Ccys = sSortedArray(sArrayRange(Ccys, MatchIDs), 2)
48        Ccys = sSubArray(Ccys, 1, 1, , 1)
49        CalculateRequiredMarketData = Ccys

50        Exit Function
ErrHandler:
51        Throw "#CalculateRequiredMarketData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreditsRequired
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : Returns a list of the credit curves that must be in the market in order to
'             be able to calculate CVA
'---------------------------------------------------------------------------------------
Function CreditsRequired()
          Dim BanksUserHasChosen
          Dim CTARes
          Dim NumTrades As Long
          Dim TR As Range

1         On Error GoTo ErrHandler
2         Set TR = getTradesRange(NumTrades)

3         If NumTrades = 0 Then
4             CreditsRequired = Empty
5             Exit Function
6         End If

7         ChooseBanks True, BanksUserHasChosen

          'If there is a WHATIF trade then we need all the credits
8         CreditsRequired = sArrayStack(TR.Columns(gCN_Counterparty).Value, gSELF)
9         If IsNumber(sMatch(True, sIsRegMatch("^" + gWHATIF + "$", CreditsRequired))) Then
10            CreditsRequired = sArrayStack(BanksUserHasChosen, gSELF)
11            Exit Function
12        Else
13            CreditsRequired = sRemoveDuplicates(CreditsRequired)
14            CreditsRequired = sMChoose(CreditsRequired, sIsRegMatch("^((?!^" + gWHATIF + "$).)*$", CreditsRequired))
15            CTARes = sCompareTwoArrays(CreditsRequired, sArrayStack(gSELF, BanksUserHasChosen), "Common")
16            If sNRows(CTARes) > 1 Then
17                CreditsRequired = sSubArray(CTARes, 2)
18            Else
19                CreditsRequired = FirstElementOf(BanksUserHasChosen)
20            End If

21        End If

22        Exit Function
ErrHandler:
23        Throw "#CreditsRequired (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenLinesWorkbook
' Author    : Philip Swannell
' Date      : 28-Apr-2015
' Purpose   : Returns a workbook object representing the source workbook whose
'             Fullname is given on the main sheet of this workbook at the range "LinesWorkbook"
'---------------------------------------------------------------------------------------
Function OpenLinesWorkbook(Optional HideOnOpening As Boolean, Optional ActivateOnOpening As Boolean = False) As Workbook
          Dim BookFullName As String
          Dim BookName As String
          Dim CopyOfErr As String
          Dim wb As Workbook

1         On Error GoTo ErrHandler
          Dim SUH As SolumAddin.clsScreenUpdateHandler
2         Set SUH = CreateScreenUpdateHandler()

3         BookFullName = FileFromConfig("LinesWorkbook")
4         BookName = sSplitPath(BookFullName)

5         On Error Resume Next
6         Set wb = Application.Workbooks(BookName)
7         On Error GoTo ErrHandler

8         If wb Is Nothing Then
9             If Not sFileExists(BookFullName) Then
10                Throw "Sorry, we couldn't find " + BookFullName + vbLf + "Is it possible it was moved, renamed or deleted?" + vbLf _
                      + "If so you need to change the 'Lines Workbook'" + vbLf + _
                      "setting on the Config sheet of this workbook.", True
11            End If
12            StatusBarWrap "Opening " + BookFullName
13            Set wb = Application.Workbooks.Open(BookFullName, , True)
14            StatusBarWrap False
15            If wb Is Nothing Then Throw "Cannot open " + BookFullName        'happens if "top of the call stack" is sheet calculation
16            CheckLinesWorkbook wb, False, True

17            If HideOnOpening Then
18                wb.Windows(1).Visible = False
19            End If
20        End If

21        If ActivateOnOpening Then
22            With wb.Windows(1)
23                .Visible = True
24                If .WindowState = xlMinimized Then .WindowState = xlNormal
25                .Activate
26            End With
27        Else
28            ThisWorkbook.Activate
29        End If

30        Set OpenLinesWorkbook = wb

31        Exit Function
ErrHandler:
32        CopyOfErr = "#OpenLinesWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
33        StatusBarWrap False
34        Throw CopyOfErr
End Function

