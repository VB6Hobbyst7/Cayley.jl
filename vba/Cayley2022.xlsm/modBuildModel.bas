Attribute VB_Name = "modBuildModel"
Option Explicit

Function GetHedgeHorizon()

1         On Error GoTo ErrHandler
2         GetHedgeHorizon = RangeFromSheet(shConfig, "HedgeHorizon", True, False, False, False, False)
3         If GetHedgeHorizon < 5 Or GetHedgeHorizon > 10 Then
4             Throw "HedgeHorizon on the Config sheet must be between 5 and 10"
5         End If

6         Exit Function
ErrHandler:
7         Throw "#GetHedgeHorizon (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub Test_BuildModelsInJulia()

1         On Error GoTo ErrHandler
2         JuliaLaunchForCayley
3         BuildModelsInJulia True, 1.1, 1.1

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#Test_BuildModelsInJulia (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BuildModelsInJulia
' Author    : Philip Swannell
' Date      : 20-Dec-2021
' Purpose   : Assumes JuliaLaunchForCayley has already been run
' -----------------------------------------------------------------------------------------------------------------------
Function BuildModelsInJulia(ForceRecreate As Boolean, FxShock As Double, FxVolShock As Double)

          'Need to have tolerance because of floating point loss of accuracy when round-tripping to Julia via JuliaExcel
          Const ABSTOL = 0.00000000000001
          Dim AllCcys As Variant
          Dim CCysToBuild As Variant
          Dim ErrorString As String
          Dim Expression As String
          Dim FileName As String
          Dim FileNameH As String
          Dim FileNameHX As String
          Dim FileNameX As String
          Dim JuliaResult As Variant
          Dim Message As String
          
          Dim MarketWB As Workbook
          Dim Numeraire As String

1         On Error GoTo ErrHandler

2         If Not ForceRecreate Then
3             If Not gModel_CM Is Nothing Then
4                 If Not gModel_CMS Is Nothing Then
5                     If Not gModel_CMH Is Nothing Then
6                         If Not gModel_CMHS Is Nothing Then
7                             If Abs(FxShock - gModel_CMS("fxshock")) < ABSTOL Then
8                                 If Abs(FxVolShock - gModel_CMS("fxvolshock")) < ABSTOL Then
9                                     If Abs(FxShock - gModel_CMHS("fxshock")) < ABSTOL Then
10                                        If Abs(FxVolShock - gModel_CMHS("fxvolshock")) < ABSTOL Then
11                                            Exit Function
12                                        End If
13                                    End If
14                                End If
15                            End If
16                        End If
17                    End If
18                End If
19            End If
20        End If

21        FileName = LocalTemp & "CayleyMarket.json"
22        FileNameH = LocalTemp & "CayleyMarketHistoricFxVol.json"
23        FileNameX = MorphSlashes(FileName, UseLinux())
24        FileNameHX = MorphSlashes(FileNameH, UseLinux())

          Dim ExportFile As Boolean
25        If ForceRecreate Then
26            ExportFile = True
27        ElseIf gModel_CM Is Nothing Then
28            ExportFile = True
29        ElseIf Not sFileExists(FileName) Then
30            ExportFile = True
31        ElseIf Not sFileExists(FileNameH) Then
32            ExportFile = True
33        End If

34        If ExportFile Then

35            Set MarketWB = OpenMarketWorkbook(True, False)
36            Numeraire = NumeraireFromMDWB()

37            AllCcys = AllCurrenciesInTradesWorkBook(ForceRecreate)
38            If LCase(RangeFromSheet(shConfig, "CurrenciesToInclude")) = "all" Then
39                CCysToBuild = AllCcys
40            Else
41                CCysToBuild = sCompareTwoArrays(AllCcys, _
                      sTokeniseString(RangeFromSheet(shConfig, "CurrenciesToInclude")), _
                      "Common")
42                If sNRows(CCysToBuild) = 1 Then
43                    CCysToBuild = Numeraire
44                Else
45                    CCysToBuild = sRemoveDuplicates(sArrayStack(Numeraire, sDrop(CCysToBuild, 1)))
46                End If
47            End If

48            CCysToBuild = SortCurrencies(CCysToBuild, Numeraire)
49            Message = "Building Hull White model with currencies " & sConcatenateStrings(CCysToBuild, ", ")

              'Save market data to file...
50            MessageLogWrite "Exporting market data for " & sConcatenateStrings(CCysToBuild, ", ")

51            ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
                  MarketWB, FileName, CCysToBuild, Numeraire, , 2)
52            ThrowIfError Application.Run("'" & MarketWB.FullName & "'!SaveDataFromMarketWorkbookToFile", _
                  MarketWB, FileNameH, CCysToBuild, Numeraire, , 3)
53        Else
54            Message = "Shocking Hull White model, fxshock = " & CStr(FxShock) & " fxvolshock = " & CStr(FxVolShock)
55        End If

          'Get Julia to do the work...
56        Expression = "using Cayley;" & _
              MN_CM & "," & MN_CMS & "," & MN_CMH & "," & MN_CMHS & _
              " = Cayley.cayleybuildmodels(""" & FileNameX & """,""" & FileNameHX & """," & _
              CStr(FxShock) & "," & CStr(FxVolShock) & ");" & _
              "Cayley.pack4models(" & MN_CM & "," & MN_CMS & "," & MN_CMH & "," & _
              MN_CMHS & "," & CStr(FxShock) & "," & CStr(FxVolShock) & ")"

57        MessageLogWrite Message

58        Assign JuliaResult, JuliaEvalVBA(Expression)
59        If VarType(JuliaResult) = vbString Then
60            Throw JuliaResult
61        ElseIf VarType(JuliaResult) = vbObject Then
62            Set gModel_CM = JuliaResult(MN_CM)
63            Set gModel_CMS = JuliaResult(MN_CMS)
64            Set gModel_CMH = JuliaResult(MN_CMH)
65            Set gModel_CMHS = JuliaResult(MN_CMHS)
66        End If

67        Set gMarketData = ParseJsonFile(FileName)

68        Set MarketWB = Nothing

69        Exit Function
ErrHandler:
70        ErrorString = "#BuildModelsInJulia (line " & CStr(Erl) & "): " & Err.Description & "!"
71        Throw ErrorString
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortCurrencies
' Author    : Philip Swannell
' Date      : 19-Sep-2016
' Purpose   : Helps alleviate the "add a currency and PFEs of existing trades changes" problem. Numeraire always goes
'             first though.
' -----------------------------------------------------------------------------------------------------------------------
Function SortCurrencies(CcyList, Numeraire As String)
          Const PreferenceOrder = "USD,EUR,GBP,JPY,CAD,CHF"
          Dim MatchIDs
          Dim Result
1         On Error GoTo ErrHandler
2         MatchIDs = sMatch(CcyList, sTokeniseString(Numeraire & "," & PreferenceOrder))
3         Result = sSortedArray(sArrayRange(CcyList, MatchIDs), 2)
4         Result = sSubArray(Result, 1, 1, , 1)
5         SortCurrencies = Result
6         Exit Function
ErrHandler:
7         Throw "#SortCurrencies (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Author     : Philip Swannell
' Date       : 07-Feb-2018
' Purpose    : Return a writable directory for saving results files to be communicated to Julia. Return is
'              terminated with backslash.
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp(Optional Refresh As Boolean = False)
          Static Res As String
1         On Error GoTo ErrHandler
2         If Not Refresh And Res <> "" Then
3             LocalTemp = Res
4             Exit Function
5         End If
6         Res = sEnvironmentVariable("temp") & "\@Cayley"
7         ThrowIfError sCreateFolder(Res)
8         If Not sFolderIsWritable(Res) Then Throw "Cannot create writable folder at " & Res

9         If Right(Res, 1) <> "\" Then Res = Res & "\"
10        LocalTemp = Res
11        Exit Function
ErrHandler:
12        Throw "#LocalTemp (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

