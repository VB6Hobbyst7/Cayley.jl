Attribute VB_Name = "modUtils2"
Option Explicit
Public gBlockCalculateEvent As Boolean
Public gBlockChangeEvent As Boolean
'PGS 13-11-20 Want to gradually move from "SCRiPT" branding to "XVA" branding, but without breaking things since that rebranding is only a nice to have.
Public Const gProjectName As String = "XVAFrontEnd"

'Want to change the format below? See method FixNumberFormatting that can apply number formatting across all sheets
Public Const NF_Comma0dp = "#,##0;[Red]-#,##0"
Public Const NF_Date = "dd-mmm-yyyy"
Public Const NF_Fx = "[>=100]#,##0.00;[>=10]#,##0.000;#,##0.0000"    'Show 4 decimal places, or if >10 show 3 dp, or if >=100 show 2dp - we could use conditional fomatting for more control...

Public Const Colour_LightGrey = 14277081
Public Const Colour_BlueText = 13395456
Public Const Colour_GreyText = 8421504
Public Const Colour_LightGreyText = 12566463
Public Const Colour_LightYellow = 10092543

Function UseLinux()
1         On Error GoTo ErrHandler
2         UseLinux = RangeFromSheet(shConfig, "UseLinux", False, False, True, False, False).Value = True
3         Exit Function
ErrHandler:
4         Throw "#UseLinux (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MsgBoxTitle
' Author    : Philip Swannell
' Date      : 19-Nov-2015
' Purpose   : Consistent title in message boxes
'---------------------------------------------------------------------------------------
Function MsgBoxTitle()
1         MsgBoxTitle = gProjectName
End Function

'---------------------------------------------------------------------------------------
' Procedure : CurrenciesSupported
' Author    : Philip Swannell
' Date      : 23 March 2016,
' Purpose   : Single place where we encode the list of supported currencies, data from sheet in SolumAddin
'---------------------------------------------------------------------------------------
Function CurrenciesSupported(LongForm As Boolean, AndInMDWb As Boolean)
1         On Error GoTo ErrHandler
          Dim AllCurrencies
          Dim AllLongForms
          Dim ChooseVector

2         ChooseVector = RangeFromSheet(shSAIStaticData, "AllCurrencies").Columns(1).Value
3         AllCurrencies = RangeFromSheet(shSAIStaticData, "AllCurrencies").Columns(2).Value
4         AllLongForms = RangeFromSheet(shSAIStaticData, "AllCurrencies").Columns(3).Value

5         If AndInMDWb Then
6             ChooseVector = sArrayAnd(ChooseVector, sArrayIsNumber(sMatch(AllCurrencies, CurrenciesInMarketWorkbook(False))))
7         End If

8         If LongForm Then
9             CurrenciesSupported = sArrayConcatenate(sMChoose(AllCurrencies, ChooseVector), " - ", _
                  sMChoose(AllLongForms, ChooseVector))
10        Else
11            CurrenciesSupported = sMChoose(AllCurrencies, ChooseVector)
12        End If
13        Exit Function
ErrHandler:
14        Throw "#CurrenciesSupported (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SupportedBDCs
' Author    : Philip Swannell
' Date      : 09-May-2016
' Purpose   : Returns a list of the supported BusinessDayConventions. This function needs
'             to be kept in synch with the Julia function adjustdate
'       CHANGING THIS FUNCTION? Then also make equivalent change to method ValidateBDC
'---------------------------------------------------------------------------------------
Function SupportedBDCs()
          Dim Res() As String
1         ReDim Res(1 To 5, 1 To 1)
2         Res(1, 1) = "Mod Foll"
3         Res(2, 1) = "Foll"
4         Res(3, 1) = "Mod Prec"
5         Res(4, 1) = "Prec"
6         Res(5, 1) = "None"
7         SupportedBDCs = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SupportedIRLegTypes
' Author     : Philip Swannell
' Date       : 16-Nov-2020
' Purpose    :                   Changing this function? then also change ValidateIRLegType
' -----------------------------------------------------------------------------------------------------------------------
Function SupportedIRLegTypes()
          Dim Res() As String
1         ReDim Res(1 To 3, 1 To 1)
2         Res(1, 1) = "Fixed"
3         Res(2, 1) = "IBOR"
4         Res(3, 1) = "RFR"
5         SupportedIRLegTypes = Res
End Function

'---------------------------------------------------------------------------------------
' Procedure : FixValidationForSupportedCurrencies
' Author    : Philip Swannell
' Date      : 13-Apr-2016
' Purpose   : If we change the supported currencies then data validation needs to change
'             on the Portfolio sheet
'---------------------------------------------------------------------------------------
Sub FixValidationForSupportedCurrencies()
1         FormatTradeTemplates
2         FormatTradesRange
3         Exit Sub
ErrHandler:
4         Throw "#FixValidationForSupportedCurrencies (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ModelExists
' Author    : Philip Swannell
' Date      : 13-Apr-2016
' Purpose   : Have we already created a model in Julia?
'---------------------------------------------------------------------------------------
Function ModelExists() As Boolean
1         On Error GoTo ErrHandler
2         If Not gResults Is Nothing Then
3             ModelExists = sFileExists(LocalTemp() & "Model.jls")
4         End If

5         Exit Function
ErrHandler:
6         ModelExists = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : CurrenciesInMarketWorkbook
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : Returns a list of he currencies for which there is a sheet in the market
'             data workbook. List includes inflation indices if argument WithInflation is True.
'---------------------------------------------------------------------------------------
Function CurrenciesInMarketWorkbook(WithInfation As Boolean)
          Dim wb As Workbook
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set wb = OpenMarketWorkbook()
          Dim i As Long
          Dim Res As Variant

3         Res = sReshape("", wb.Worksheets.Count, 1)
4         i = 0
5         For Each ws In wb.Worksheets
6             If IsCurrencySheet(ws) Then
7                 i = i + 1
8                 Res(i, 1) = ws.Name
9             ElseIf WithInfation Then
10                If IsInflationSheet(ws) Then
11                    i = i + 1
12                    Res(i, 1) = ws.Name
13                End If
14            End If

15        Next ws
16        Res = sSubArray(Res, 1, 1, i)
17        CurrenciesInMarketWorkbook = Res
18        Exit Function
ErrHandler:
19        Throw "#CurrenciesInMarketWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreditsInModel
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : Returns a list of the names for which credit curves have been built
'---------------------------------------------------------------------------------------
Function CreditsInModel()
1         On Error GoTo ErrHandler

2         CreditsInModel = Empty
3         On Error Resume Next
4         CreditsInModel = sArrayTranspose(gResults("Model")("CreditCurve").keys)
5         Exit Function
ErrHandler:
6         Throw "#CreditsInModel (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CurrenciesInModel
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : Returns a list of the currencies for which curves have been built
'---------------------------------------------------------------------------------------
Function CurrenciesInModel()
1         On Error GoTo ErrHandler

2         CurrenciesInModel = Empty
3         On Error Resume Next
4         CurrenciesInModel = sArrayTranspose(gResults("Model")("Currencies"))
5         Exit Function
ErrHandler:
6         Throw "#CurrenciesInModel (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : TestCheckModelHasCurves
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : test harness
'---------------------------------------------------------------------------------------
Sub TestCheckModelHasCurves()
          Dim Ccys
          Dim IIs
          Dim Res
1         On Error GoTo ErrHandler
2         CalculateRequiredMarketData Ccys, IIs

3         Res = CheckModelHasCurves(Ccys, CreditsRequired())
4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TestCheckModelHasCurves (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckModelHasCurves
' Author    : Philip Swannell
' Date      : 19-May-2016
' Purpose   : Returns True if both:
'          a) CcysReq (a list of ISO currency codes) are already in the market
'             (there is an implicit assumption that if the discount curve exists then so
'             too does the calibrated HW sigmas and the necessary Fx spot and vol levels);
'          b) CrdsRequired are already credit curves in the market
'             Returns False otherwise, but instead throws an error if the "missing" data in the market
'             is not available in the market data workbook.
'             In this way method XVAFrontEndMain can automatically rebuild the model when necessary
'             and if required market data is missing then this method throws a friendlier error
'             message than would be thrown by low-level Julia functions.
'---------------------------------------------------------------------------------------
Function CheckModelHasCurves(Optional ByVal CcysReq As Variant, Optional CrdsRequired, Optional IIsReq)
          Dim CCysInModel
          Dim CCysInMWB
          Dim CcysMissing
          Dim ChooseVector
          Dim CrdsInModel

1         On Error GoTo ErrHandler
2         If IsEmpty(CcysReq) Or IsMissing(CcysReq) Then
3             CalculateRequiredMarketData CcysReq, IIsReq
4         End If

5         If Not IsMissing(IIsReq) Then    'because in the HW model, inflation indices are currencies
6             CcysReq = sArrayStack(CcysReq, IIsReq)
7         End If

8         CCysInModel = CurrenciesInModel()

9         ChooseVector = sArrayIsNumber(sMatch(CcysReq, CCysInModel))
10        CheckModelHasCurves = True

11        If Not sColumnAnd(ChooseVector)(1, 1) Then
12            CcysMissing = sMChoose(CcysReq, sArrayNot(ChooseVector))
13            CCysInMWB = CurrenciesInMarketWorkbook(True)
14            ChooseVector = sArrayIsText(sMatch(CcysMissing, CCysInMWB))
15            If sColumnOr(ChooseVector)(1, 1) Then
16                Throw "The trades require interest rate and volatility curves for " + sConcatenateStrings(sMChoose(CcysMissing, ChooseVector), ", ") + " but these are not available in the Market Data Workbook"
17            Else
18                CheckModelHasCurves = False
19            End If
20        End If

21        If Not (IsMissing(CrdsRequired) Or IsEmpty(CrdsRequired)) Then
22            CrdsInModel = CreditsInModel()
23            ChooseVector = sArrayIsNumber(sMatch(CrdsRequired, CreditsInModel))
24            If Not sColumnAnd(ChooseVector)(1, 1) Then
25                CheckModelHasCurves = False
26            End If
27        End If

28        Exit Function
ErrHandler:
29        Throw "#CheckModelHasCurves (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : RangeFromMarketDataBook
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : We want to get values from the Market Data workbook without setting up the dreaded Excel links
'---------------------------------------------------------------------------------------
Function RangeFromMarketDataBook(SheetName As String, ByVal RangeName As String)
1         Application.Volatile

          Dim BookFullName As String
          Dim BookName As String
          Dim ErrString As String
          Dim N As Name
          Dim wb As Workbook
          Dim ws As Worksheet

2         On Error GoTo ErrHandler
3         BookFullName = FileFromConfig("MarketDataWorkbook")
4         BookName = sSplitPath(BookFullName)

5         On Error Resume Next
6         Set RangeFromMarketDataBook = Application.Workbooks(BookName).Worksheets(SheetName).Names(RangeName).RefersToRange
7         If Err.Number = 0 Then Exit Function
8         On Error GoTo ErrHandler

          'Most likely the problem is that the market data workbook is not open, so open it, though this won't work if caller is an cell formula.
9         OpenMarketWorkbook True, False
10        On Error Resume Next
11        Set RangeFromMarketDataBook = Application.Workbooks(BookName).Worksheets(SheetName).Names(RangeName).RefersToRange
12        If Err.Number = 0 Then Exit Function
13        On Error GoTo ErrHandler

          'Still didn't work so construct a friendly error message
14        If Not IsInCollection(Application.Workbooks, BookName) Then Throw "Market Data Workbook (" + BookName + ") is not open"
15        Set wb = Application.Workbooks(BookName)
16        If Not IsInCollection(wb.Worksheets, SheetName) Then Throw "#Cannot find worksheet '" + SheetName + "' in MarketDataWorkbook"
17        Set ws = wb.Worksheets(SheetName)
18        If Not IsInCollection(ws.Names, RangeName) Then
              'PGS 28 March 2017. Changed "BaseCCY" to "Numeraire" on Config sheet of MDW
19            If SheetName = "Config" And RangeName = "BaseCCY" And IsInCollection(ws.Names, "Numeraire") Then
20                RangeName = "Numeraire"
21            Else
22                Throw "#Cannot find range named '" + RangeName + "' in sheet '" + SheetName + "' of MarketDataWorkbook"
23            End If
24        End If
25        Set N = ws.Names(RangeName)
26        If Not NameRefersToRange(N) Then Throw "Name '" + RangeName + "' on sheet '" + SheetName + "' of Market Data book does not refere to a range"
27        Throw "Unknown Error"

28        Exit Function
ErrHandler:
29        ErrString = "#RangeFromMarketDataBook (line " & CStr(Erl) + "): " & Err.Description & "!"
30        If TypeName(Application.Caller) = "Range" Then
31            RangeFromMarketDataBook = ErrString
32        Else
33            Throw ErrString
34        End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : NameRefersToRange
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Returns TRUE if the name refers to a Range, FALSE if it referes to something
'             else.
'---------------------------------------------------------------------------------------
Private Function NameRefersToRange(TheName As Name) As Boolean
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = TheName.RefersToRange
3         NameRefersToRange = True
4         Exit Function
ErrHandler:
5         NameRefersToRange = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : FixNumberFormatting
' Author    : Philip Swannell
' Date      : 24-May-2016
' Purpose   : Make number formatting consistent on all sheets
'---------------------------------------------------------------------------------------
Sub FixNumberFormatting()
          Dim c As Range
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim ws As Worksheet
1         For Each ws In ThisWorkbook.Worksheets
2             Set SPH = CreateSheetProtectionHandler(ws)
3             For Each c In ws.UsedRange.Cells
4                 If InStr(c.NumberFormat, "[Red]") > 0 Then
5                     If IsEmpty(c.Value) Then
6                         c.NumberFormat = "General"
7                     ElseIf c.NumberFormat <> NF_Comma0dp Then
8                         c.NumberFormat = NF_Comma0dp
9                     End If
10                End If
11            Next c
12        Next ws
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UncompressTradesOnPortfolioSheet
' Author    : Philip Swannell
' Date      : 09-Jan-2017
' Purpose   : Available from menu...
'---------------------------------------------------------------------------------------
Sub UncompressTradesOnPortfolioSheet()
          Dim ExistingTrades As Range
          Dim JuliaTrades
          Dim MatchRes1
          Dim MatchRes2
          Dim NewTrades
          Dim Numeraire As String
          Dim NumTrades As Long
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim VFs
1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         OpenMarketWorkbook True, False
4         Set ExistingTrades = getTradesRange(NumTrades)
5         If NumTrades = 0 Then Exit Sub
6         VFs = ExistingTrades.Columns(gCN_TradeType).Value
7         MatchRes1 = sMatch("FxOptionStrip", VFs)
8         MatchRes2 = sMatch("FxForwardStrip", VFs)
9         If Not (IsNumber(MatchRes1) Or IsNumber(MatchRes2)) Then Exit Sub
10        Numeraire = RangeFromMarketDataBook("Config", "Numeraire")
11        JuliaTrades = PortfolioTradesToJuliaTrades(ExistingTrades.Value2, True, False)
12        JuliaTrades = UncompressTrades(JuliaTrades)
13        NewTrades = JuliaTradesToPortfolioTrades(JuliaTrades, Numeraire)
14        PasteTradesToPortfolioSheet NewTrades, RangeFromSheet(shPortfolio, "TradesFileName"), True
15        Exit Sub
ErrHandler:
16        Throw "#UncompressTradesOnPortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CompressTradesOnPortfolioSheet
' Author    : Philip Swannell
' Date      : 09-Jan-2017
' Purpose   : Available from menu...
'---------------------------------------------------------------------------------------
Sub CompressTradesOnPortfolioSheet()
          Dim ExistingTrades As Range
          Dim JuliaTrades
          Dim NewTrades
          Dim NumCompressedTrades As Long
          Dim Numeraire As String
          Dim NumTrades As Long
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim VFs
1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         OpenMarketWorkbook True, False
4         Set ExistingTrades = getTradesRange(NumTrades)
5         If NumTrades = 0 Then Exit Sub
6         VFs = ExistingTrades.Columns(gCN_TradeType).Value
7         Numeraire = RangeFromMarketDataBook("Config", "Numeraire")
8         JuliaTrades = PortfolioTradesToJuliaTrades(ExistingTrades.Value2, True, False)
9         JuliaTrades = UncompressTrades(JuliaTrades)    'We have to first uncompress then recompress to merge together (e.g.) FxForward and any existing FxForwardStrip
10        JuliaTrades = CompressFxForwards(JuliaTrades, NumCompressedTrades)
11        JuliaTrades = CompressFxOptions(JuliaTrades, NumCompressedTrades + 1)
12        NewTrades = JuliaTradesToPortfolioTrades(JuliaTrades, Numeraire)
13        PasteTradesToPortfolioSheet NewTrades, RangeFromSheet(shPortfolio, "TradesFileName"), True
14        Exit Sub
ErrHandler:
15        Throw "#CompressTradesOnPortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConfigRange
' Author     : Philip Swannell
' Date       : 17-Aug-2020
' Purpose    : Get info from the Config sheet of this workbook, handles the fact that merged calls may mean we need to
'              use the .cells(1,1) property.
' PGS 23 March 2022. No longer have merged cells on the Config sheet - more trouble than they're worth
' -----------------------------------------------------------------------------------------------------------------------
Function ConfigRange(Name As String) As Range
1         Set ConfigRange = RangeFromSheet(shConfig, Name).Cells(1, 1)
End Function

'---------------------------------------------------------------------------------------
' Procedure : LookupBankInfo
' Author    : Philip Swannell
' Date      : 11-Jan-2017
' Purpose   : Do a lookup into the data on the lines workbook without establishing an Excel link to the lines workbook
'             BankName can be a column or string and ParameterName can be column or string, but they cannot both be columns
'---------------------------------------------------------------------------------------
Function LookupBankInfo(ByVal BankName As Variant, ByVal ParameterName As Variant)
1         Application.Volatile
          Dim Counterparties
          Dim Headers
          Dim i As Long
          Dim lo As ListObject
          Dim MatchID1
          Dim MatchIDsB
          Dim MatchIDsP
          Dim NBanks As Long
          Dim NParams As Long
          Dim Result()
          Dim shCapInputs As Worksheet
          Const BankNameHeader = "CPTY_PARENT"

2         On Error GoTo ErrHandler
3         If sNCols(ParameterName) > 1 Then Throw "ParameterName must be a string or column array of strings"
4         If sNCols(BankName) > 1 Then Throw "BankName must be a string or column array of strings"
5         NParams = sNRows(ParameterName)
6         NBanks = sNRows(BankName)
7         If NParams > 1 Then If NBanks > 1 Then Throw "BankName and ParameterName cannot both have more than one element"

8         On Error Resume Next
9         Set shCapInputs = Application.Workbooks(sSplitPath(ConfigRange("LinesWorkbook").Value)).Worksheets(SN_Lines)
10        On Error GoTo ErrHandler

11        If shCapInputs Is Nothing Then Throw "Lines workbook is not open"
12        Select Case shCapInputs.ListObjects.Count
              Case 0
13                Throw "Cannot find Table on sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'"
14            Case Is > 1
15                Throw "Unexpected error: more than one table on sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'"
16        End Select

17        Set lo = shCapInputs.ListObjects(1)
18        Headers = sArrayTranspose(lo.HeaderRowRange)
19        MatchID1 = sMatch(BankNameHeader, Headers)
20        If Not IsNumber(MatchID1) Then Throw "Cannot find header '" + BankNameHeader + "' on Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'"
21        MatchIDsP = sMatch(ParameterName, Headers)
22        If NParams = 1 Then If Not IsNumber(MatchIDsP) Then Throw "Cannot find header '" + ParameterName + "' on Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'"
23        Counterparties = lo.DataBodyRange.Columns(MatchID1).Value
24        MatchIDsB = sMatch(BankName, Counterparties)

25        If NParams = 1 And NBanks = 1 Then
26            LookupBankInfo = lo.DataBodyRange.Cells(MatchIDsB, MatchIDsP).Value
27        ElseIf NParams > 1 And NBanks = 1 Then
28            Force2DArrayR ParameterName
29            ReDim Result(1 To NParams, 1 To 1)
30            For i = 1 To NParams
31                If Not IsNumber(MatchIDsB) Then
32                    Result(i, 1) = "#Cannot find bank '" + CStr(BankName) + "' on Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'!"
33                ElseIf Not IsNumber(MatchIDsP(i, 1)) Then
34                    Result(i, 1) = "#Cannot find header '" + CStr(ParameterName(i, 1)) + "' on Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'!"
35                Else
36                    Result(i, 1) = lo.DataBodyRange.Cells(MatchIDsB, MatchIDsP(i, 1)).Value
37                End If
38            Next i
39            LookupBankInfo = Result
40        ElseIf NBanks > 1 And NParams = 1 Then
41            Force2DArrayR BankName
42            ReDim Result(1 To NBanks, 1 To 1)
43            For i = 1 To NBanks
44                If Not IsNumber(MatchIDsP) Then
45                    Result(i, 1) = "#Cannot find header '" + CStr(FirstElementOf(ParameterName)) + "' on Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'!"
46                ElseIf Not IsNumber(MatchIDsB(i, 1)) Then
47                    Result(i, 1) = "#Cannot find bank '" + CStr(BankName(i, 1)) + "' in column'" + BankNameHeader + "' of Table1 in sheet '" + SN_Lines + "' of workbook '" + shCapInputs.Parent.Name + "'!"
48                Else
49                    Result(i, 1) = lo.DataBodyRange.Cells(MatchIDsB(i, 1), MatchIDsP).Value
50                End If
51            Next i
52            LookupBankInfo = Result
53        End If

54        Exit Function
ErrHandler:
55        LookupBankInfo = "#LookupBankInfo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Run_xva_main
' Author     : Philip Swannell
' Date       : 12-Nov-2020
' Purpose    : Call Julia method xva_main in a running instance of the Julia REPL, using JuliaExcel https://github.com/PGS62/JuliaExcel.jl
' -----------------------------------------------------------------------------------------------------------------------
Function Run_xva_main(ByVal ControlFile As String, ResultsFileFullName As String, BuildModelFromDFsAndSurvProbs As Boolean) As Dictionary
          Dim CodeToExecute As String
          Dim ControlFile2 As String
          Dim XSH As clsExcelStateHandler
          Const DQ = """"
          Dim CommandLineOptions As String
          Dim Prompt As String
          Dim Res As Variant
          Dim SystemImage As String
          Dim SystemImageX As String

1         On Error GoTo ErrHandler

2         SystemImage = IIf(UseLinux, gSysImageXVALinux, gSysImageXVAWindows)
3         SystemImageX = MorphSlashes(SystemImage, UseLinux())
4         Set XSH = CreateExcelStateHandler(, , , "Waiting for Julia code to execute")

5         If Not sFileExists(ControlFile) Then Throw "Cannot find file '" + ControlFile + "'"
6         If sFileExists(ResultsFileFullName) Then
7             ThrowIfError sFileDelete(ResultsFileFullName)
8         End If

9         If Not JuliaIsRunning Then
10            If sFileExists(SystemImage) Then
11                CommandLineOptions = " --threads auto --sysimage " & SystemImageX
12            Else
13                Prompt = "There is no ""system image"" available. What would you like to do?" + vbLf + vbLf + _
                      "Recommended: Create a system image now. This will take about 10 minutes and but will save time in the future, by doing ""ahead-of-time"" compilation of the Julia code." + vbLf + vbLf + _
                      "Not recommended, except for developer use: Run Julia without a system image."
14                Select Case MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, MsgBoxTitle(), "Create a system image now", "Run Julia without system image")
                      Case vbYes
                          JuliaCreateSystemImage True, UseLinux()
15                        Exit Function
16                    Case vbNo
17                        CommandLineOptions = " --threads auto"
18                    Case vbCancel
19                        Throw "User Cancelled", True
20                End Select
21            End If
22            ThrowIfError julialaunch(UseLinux, False, CommandLineOptions, "XVA")
23        End If
              
24        ControlFile = MorphSlashes(ControlFile, UseLinux())
25        If BuildModelFromDFsAndSurvProbs Then
26            If LCase(Right(ControlFile, 5)) = ".json" Then
27                ControlFile2 = Left(ControlFile, Len(ControlFile) - 5) + "2" + Right(ControlFile, 5)
28            Else
29                ControlFile2 = ControlFile + "2"
30            End If
31        Else
32            ControlFile2 = ControlFile
33        End If
              
34        CodeToExecute = "xva_main(" + DQ + ControlFile2 + DQ + "," + LCase(sEquals(ConfigRange("UseThreads"), True)) + ")"

35        If BuildModelFromDFsAndSurvProbs Then
36            CodeToExecute = "XVA.transformfiles(" + DQ + ControlFile + DQ + "," + DQ + ControlFile2 + DQ + ");" & CodeToExecute
37        End If

38        CodeToExecute = "using Revise;using XVA; " & CodeToExecute
39        assign Res, JuliaEvalVBA(CodeToExecute)
40        ThrowIfError sFileCopy(JuliaExcel.JuliaResultFile(), sJoinPath(LocalTemp(), LocalTemp() & "results.txt"))
        
41        If VarType(Res) = vbString Then Throw Res
        
42        If TypeName(Res) <> "Dictionary" Then
43            Throw "Expected data returned from Julia to be a Dictionary, but got data of type " + TypeName(Res)
44        End If

45        Set Run_xva_main = Res

46        Exit Function
ErrHandler:
47        Throw "#Run_xva_main (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub LaunchJuliaWithoutSystemImage()

1         On Error GoTo ErrHandler
2         If JuliaIsRunning Then
3             JuliaEval "exit()"
4         End If
5         ThrowIfError julialaunch(UseLinux, False, " --threads=auto", "XVA")
6         Exit Sub
ErrHandler:
7         Throw "#LaunchJuliaWithoutSystemImage (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

