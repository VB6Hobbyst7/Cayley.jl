Attribute VB_Name = "modGenerateTestTrades"
Option Explicit

Sub CreateTestSet()

          Const ProjectName = "c:\Projects\XVA\"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Prompt
          Dim TargetFolder As String
          Const FileNames As String = "Control.json,MarketDiscountFactors.json,MarketRates.json,Model.jls,Results.json,Trades.csv"
          Dim Res As VbMsgBoxResult
          Dim SourceFiles
          Dim TargetFiles
          Static ReadMeText As String
          Dim NumTrades As Long
          
          'Too easy to make a mistake - check trades not changed since last run!
          Dim TradesOnDisk
          Dim TradesOnPortfolio
          
1         On Error GoTo ErrHandler
2         TradesOnDisk = sFileShow(LocalTemp + "Trades.csv", ",", True, True, True, , "yyyy-mm-dd")
3         TradesOnPortfolio = getTradesRange(NumTrades, False).Value2
4         If NumTrades = 0 Then Throw "For a test set to be created there must be trades on the Portfolio sheet.", True
5         TradesOnPortfolio = PortfolioTradesToJuliaTrades(TradesOnPortfolio, True, False)
          
6         Force2DArrayR TradesOnPortfolio, NR, NC
7         For i = 1 To NR
8             For j = 1 To NC
9                 If IsEmpty(TradesOnPortfolio(i, j)) Then
10                    TradesOnPortfolio(i, j) = ""
11                End If
12            Next
13        Next

14        If Not sArraysNearlyIdentical(TradesOnDisk, TradesOnPortfolio, True, 0.00000001) Then
15            g sArrayTranspose(TradesOnDisk)
16            g sArrayTranspose(TradesOnPortfolio)
17            g sArrayTranspose(sDiffTwoArrays(TradesOnDisk, TradesOnPortfolio))
          
18            Throw "Trades have changed since the last valuation"
19        End If

20        On Error GoTo ErrHandler
21        i = 1
22        While sFileExists(ProjectName + "data\set" & CStr(i) & "\Control.json")
23            i = i + 1
24        Wend
25        TargetFolder = ProjectName + "data\set" & CStr(i)

26        Prompt = "Copy files to duplicate the most recent valuation, which had the following parameters:" + vbLf + sConcatenateStrings(sFileShow(LocalTemp + "Control.json", ""), vbLf)

27        Res = MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, MsgBoxTitle())
28        If Res <> vbOK Then Exit Sub

29        Prompt = "Create folder " + TargetFolder + "'?"
30        Res = MsgBoxPlus(Prompt, vbQuestion + vbYesNoCancel, MsgBoxTitle(), "Yes", "No, Choose another folder", "Quit")
31        If Res = vbCancel Then
32            Exit Sub
33        ElseIf Res = vbNo Then
34            TargetFolder = FolderPicker(ProjectName, , , "SaveXVAFilesToTestSets", True)
35            If TargetFolder = "False" Then Exit Sub
36        End If

37        SourceFiles = sJoinPath(LocalTemp(), sTokeniseString(FileNames))
38        SourceFiles = sArrayStack(SourceFiles, sJoinPath(LocalTemp(), "Results.json"))
39        TargetFiles = sJoinPath(TargetFolder, sTokeniseString(FileNames))
40        TargetFiles = sArrayStack(TargetFiles, sJoinPath(TargetFolder, "ResultsForRegression.json"))

          Dim ButtonClicked As String
TryAgain:
41        ReadMeText = InputBoxPlus("ReadMe contents", , ReadMeText, , , 400, 60, , , , , ButtonClicked)
42        If ButtonClicked = "Cancel" Then Exit Sub
43        If ReadMeText = "" Then GoTo TryAgain

          Dim f As Variant
44        For Each f In SourceFiles
45            If Not sFileExists(f) Then Throw "Cannot find file '" + f + "'"
46        Next f

          Dim NumExist As Long
47        For Each f In TargetFiles
48            If sFileExists(f) Then NumExist = NumExist + 1
49        Next

50        If NumExist > 0 Then
51            Prompt = "Copy files from '" + CStr(LocalTemp()) + "' to '" + TargetFolder + "'?"
52            Prompt = Prompt + vbLf + vbLf + "NB " + CStr(NumExist) + " of the target files already exist!"
53            Res = MsgBoxPlus(Prompt, vbOKCancel + vbExclamation, MsgBoxTitle(), "Yes, Copy", "No, Quit")
54            If Res <> vbOK Then Exit Sub
55        End If

56        ThrowIfError sCreateFolder(TargetFolder)
57        ThrowIfError sFirstError(sFileCopy(SourceFiles, TargetFiles))
          Dim ReadMeFile As String
58        ReadMeFile = sJoinPath(TargetFolder, "readme.md")
59        ThrowIfError sFileSave(ReadMeFile, ReadMeText, "")

60        Prompt = "All done, created files:" + vbLf + _
              sConcatenateStrings(sArrayStack(TargetFiles, ReadMeFile), vbLf)

61        MsgBoxPlus Prompt, vbInformation, MsgBoxTitle()

62        Exit Sub
ErrHandler:
63        SomethingWentWrong "#CreateTestSet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function MyRnd()
1         MyRnd = sRandomVariable(1)(1, 1)
End Function

Sub GenerateTestSwaps()

          Dim AllDCTs As Variant
          Dim BDC1 As Variant
          Dim BDC2 As Variant
          Dim BDCs As Variant
          Dim Ccy1 As Variant
          Dim Ccy2 As Variant
          Dim Counterparty As Variant
          Dim DCT1 As Variant
          Dim DCT2 As Variant
          Dim DCTs As Variant
          Dim EndDate As Variant
          Dim FloatingDCTs As Variant
          Dim Freq1 As Variant
          Dim Freq2 As Variant
          Dim LegType1 As Variant
          Dim LegType2 As Variant
          Dim Notional1 As Variant
          Dim Notional2 As Variant
          Dim Rate1 As Variant
          Dim Rate2 As Variant
          Dim StartDate As Variant
          Dim STK As clsStacker
          Dim Trade(1 To 1, 1 To 19) As Variant
          Dim TradeID As Variant
          Dim TradeType As Variant

          Dim i As Long

1         AllDCTs = sSupportedDCTs()
2         FloatingDCTs = sArrayStack("A/360", "A/365F", "ActB/ActB")

3         ThrowIfError sRandomSetSeed("Wichmann-Hill", 123)

4         BDCs = SupportedBDCs()
5         Set STK = CreateStacker()

6         For Each TradeType In Array("InterestRateSwap")
7             For Each StartDate In Array(DateValue("12-Feb-2020"))
8                 For Each EndDate In Array(DateValue("12-Feb-2025"))
9                     For Each Ccy1 In Array("EUR")
10                        For Each Notional1 In Array(10000000, 100.5)
11                            For Each Rate1 In Array(0.01, 0)
12                                For Each LegType1 In Array("Fixed", "Libor", "OIS")
13                                    If LegType1 = "Fixed" Then
14                                        DCTs = AllDCTs
15                                    Else
16                                        DCTs = FloatingDCTs
17                                    End If
18                                    For Each Freq1 In Array("Annual", "Quarterly")
19                                        For Each DCT1 In DCTs
20                                            For Each BDC1 In BDCs
21                                                i = i + 1
22                                                TradeID = "T" + Format(i, "000000")
23                                                Ccy2 = Ccy1
24                                                Notional2 = Notional1
25                                                Rate2 = Choose((i Mod 2) + 1, 0.01, 0)
26                                                LegType2 = Choose((i Mod 3) + 1, "Fixed", "Libor", "OIS")
27                                                Freq2 = Choose((i Mod 3) + 1, "Annual", "Semi annual", "Quarterly", "Monthly")
28                                                If LegType2 = "Fixed" Then
29                                                    DCT2 = AllDCTs((i Mod 8) + 1, 1)
30                                                Else
31                                                    DCT2 = FloatingDCTs((i Mod 3) + 1, 1)
32                                                End If
33                                                BDC2 = BDCs((i Mod 4) + 1, 1)
34                                                Counterparty = "BARC_GB_LON"
35                                                Trade(1, 1) = TradeID
36                                                Trade(1, 2) = TradeType
37                                                Trade(1, 3) = StartDate
38                                                Trade(1, 4) = EndDate
39                                                Trade(1, 5) = Ccy1
40                                                Trade(1, 6) = Notional1
41                                                Trade(1, 6) = CLng(5000 + MyRnd() * 5000) * 1000
42                                                Trade(1, 7) = Rate1
43                                                Trade(1, 8) = LegType1
44                                                Trade(1, 9) = Freq1
45                                                Trade(1, 10) = DCT1
46                                                Trade(1, 11) = BDC1
47                                                Trade(1, 12) = Ccy2
48                                                Trade(1, 13) = Notional2
49                                                Trade(1, 13) = Trade(1, 6)
50                                                Trade(1, 14) = Rate2
51                                                Trade(1, 15) = LegType2
52                                                Trade(1, 16) = Freq2
53                                                Trade(1, 17) = DCT2
54                                                Trade(1, 18) = BDC2
55                                                Trade(1, 19) = Counterparty

56                                                STK.Stack2D Trade
57                                            Next
58                                        Next
59                                    Next
60                                Next
61                            Next
62                        Next
63                    Next
64                Next
65            Next
66        Next

67        PasteTradesToPortfolioSheet STK.report, , True

End Sub

Sub GenerateTestCapFloors()

          Dim BDC1 As Variant
          Dim BDCs As Variant
          Dim Ccy1 As Variant
          Dim Counterparty As Variant
          Dim DCT1 As Variant
          Dim DCTs As Variant
          Dim EndDate As Variant
          Dim Freq1 As Variant
          Dim LegType1 As Variant
          Dim Notional1 As Variant
          Dim Rate1 As Variant
          Dim StartDate As Variant
          Dim STK As clsStacker
          Dim Trade(1 To 1, 1 To 19) As Variant
          Dim TradeID As Variant
          Dim TradeType As Variant

          Dim i As Long

1         DCTs = sArrayStack("A/360", "A/365F")

2         ThrowIfError sRandomSetSeed("Wichmann-Hill", 123)

3         BDCs = SupportedBDCs()
4         Set STK = CreateStacker()

5         For Each TradeType In Array("CapFloor")
6             For Each StartDate In Array(DateValue("12-Feb-2020"))
7                 For Each EndDate In Array(DateValue("12-Feb-2025"))
8                     For Each Ccy1 In Array("EUR")
9                         For Each Notional1 In Array(10000000, 100.5)
10                            For Each Rate1 In Array(0.01, 0)
11                                For Each LegType1 In Array("BuyCap", "SellFloor")
12                                    For Each Freq1 In Array("Semi annual", "Quarterly", "Monthly")
13                                        For Each DCT1 In DCTs
14                                            For Each BDC1 In BDCs
15                                                i = i + 1
16                                                TradeID = "T" + Format(i, "000000")
17                                                Counterparty = Choose((i Mod 3) + 1, "BARC_GB_LON", "RBOS_GB_LON", "LOYD_GB_LON")
18                                                Trade(1, 1) = TradeID
19                                                Trade(1, 2) = TradeType
20                                                Trade(1, 3) = StartDate
21                                                Trade(1, 4) = EndDate
22                                                Trade(1, 5) = Ccy1
23                                                Trade(1, 6) = Notional1
24                                                Trade(1, 6) = CLng(5000 + MyRnd() * 5000) * 1000
25                                                Trade(1, 7) = (MyRnd() - 0.5) / 100
26                                                Trade(1, 8) = LegType1
27                                                Trade(1, 9) = Freq1
28                                                Trade(1, 10) = DCT1
29                                                Trade(1, 11) = BDC1
30                                                Trade(1, 12) = CVErr(xlErrNA)
31                                                Trade(1, 13) = CVErr(xlErrNA)
32                                                Trade(1, 13) = CVErr(xlErrNA)
33                                                Trade(1, 14) = CVErr(xlErrNA)
34                                                Trade(1, 15) = CVErr(xlErrNA)
35                                                Trade(1, 16) = CVErr(xlErrNA)
36                                                Trade(1, 17) = CVErr(xlErrNA)
37                                                Trade(1, 18) = CVErr(xlErrNA)
38                                                Trade(1, 19) = Counterparty

39                                                STK.Stack2D Trade
40                                            Next
41                                        Next
42                                    Next
43                                Next
44                            Next
45                        Next
46                    Next
47                Next
48            Next
49        Next

50        PasteTradesToPortfolioSheet STK.report, , True

End Sub

