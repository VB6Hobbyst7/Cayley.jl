Attribute VB_Name = "modMarketData"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modMarketData
' Author    : Philip Swannell
' Date      : 9-Sep-2015
' Purpose   : Methods to grab market, either from the Julia environment directly, or else from "model bare bones"
'             dictionaries that were returned from Julia further up the call stack.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Base 1

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MyFxPerBaseCcy
' Author     : Philip Swannell
' Date       : 17-Jan-2022
' Purpose    : Calculate spot fx rates not by interrogating R (as previously) but by interrogating a dictionary
'              returned from Julia. Unlike its predecessor function the return has the same number of dimensions as Ccy
'              rather than always having two dimensions.
' Parameters :
'  Ccy       : a string, or an array of strings (in fact currency iso codes)
'  BaseCCY       : The
'  BareBonesModel: A dictionary, with element "spot" that in turn has elements keyed on the currency codes and giving
'                  spot rates against (an unspecified) numeraire.
' -----------------------------------------------------------------------------------------------------------------------
Function MyFxPerBaseCcy(ByVal Ccy As Variant, BaseCCY As String, BareBonesModel As Dictionary)

          Dim baserate As Double
          Dim i As Long
          Dim j As Long
          Dim Result() As Variant
          Dim spots As Dictionary

1         On Error GoTo ErrHandler
2         Set spots = DictGet(BareBonesModel, "spot")

3         If VarType(Ccy) = vbString Then
4             MyFxPerBaseCcy = DictGet(spots, CStr(Ccy)) / DictGet(spots, BaseCCY)
5         ElseIf IsArray(Ccy) Then
6             baserate = DictGet(spots, BaseCCY)
7             Select Case NumDimensions(Ccy)

                  Case 1
8                     ReDim Result(LBound(Ccy) To UBound(Ccy))
9                     For i = LBound(Ccy) To UBound(Ccy)
10                        Result(i) = DictGet(spots, CStr(Ccy(i))) / baserate
11                    Next i
12                Case 2
13                    ReDim Result(LBound(Ccy, 1) To UBound(Ccy, 1), LBound(Ccy, 2) To UBound(Ccy, 2))
14                    For i = LBound(Ccy, 1) To UBound(Ccy, 1)
15                        For j = LBound(Ccy, 2) To UBound(Ccy, 2)
16                            Result(i, j) = DictGet(spots, CStr(Ccy(i, j))) / baserate
17                        Next j
18                    Next i
19                Case Else
20                    Throw "unexpected number of dimensions in TheCCys"
21            End Select
22            MyFxPerBaseCcy = Result
23        Else
24            Throw "Ccy is of unexpected type: " & TypeName(Ccy)
25        End If

26        Exit Function
ErrHandler:
27        Throw "#MyFxPerBaseCcy (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : EURUSDForwardRates
' Author     : Philip Swannell
' Date       : 26-Jan-2022
' Purpose    : Returns an array of fx forward rates. This function calls Julia directly and is therefore slow
'              (at least until I upgrade JuliaExcel) so use sparingly. Currently used only in method ExecuteTrades.
' Parameters :
'  TheDates : Dates - two dimensions with 1 column.
'  ModelName: The names of the model in Julia
' -----------------------------------------------------------------------------------------------------------------------
Function EURUSDForwardRates(ByVal TheDates, ModelName As String)
          Dim DatesAsJuliaLiteral As String
          Dim Expression As String
          Dim i As Long

1         On Error GoTo ErrHandler
2         For i = 1 To sNRows(TheDates)
3             If Not IsNumberOrDate(TheDates(i, 1)) Then Throw "Cannot convert TheDates to a vector of Longs"
4             TheDates(i, 1) = CLng(TheDates(i, 1))
5         Next i

6         DatesAsJuliaLiteral = "[" & sConcatenateStrings(TheDates) & "]"
7         Expression = "Cayley.EURUSDforwardrates(" & ModelName & "," & DatesAsJuliaLiteral & ")"
8         EURUSDForwardRates = ThrowIfError(JuliaEvalVBA(Expression))

9         Exit Function
ErrHandler:
10        Throw "#EURUSDForwardRates (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function TenorToTimeCore(ByVal Tenor As String)
          Dim TheNumber As Double

1         On Error GoTo ErrHandler

2         TheNumber = -100
3         On Error Resume Next
4         TheNumber = CDbl(Left(CStr(Tenor), Len(CStr(Tenor)) - 1))
5         On Error GoTo ErrHandler
6         If TheNumber = -100 Then Throw "Unrecognised Tenor: " & CStr(Tenor)
7         Select Case UCase(Right(Tenor, 1))
              Case "Y"
8                 TenorToTimeCore = TheNumber
9             Case "M"
10                TenorToTimeCore = TheNumber / 12
11            Case "W"
12                TenorToTimeCore = TheNumber * 7 / 365.25
13            Case "D"
14                TenorToTimeCore = TheNumber / 365.25
15            Case Else
16                Throw "Unrecognised Tenor: " & CStr(Tenor)
17        End Select

18        Exit Function
ErrHandler:
19        Throw "#TenorToTime (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TenorToTime
' Author    : Philip Swannell
' Date      : 12-Sep-2016
' Purpose   : Replicates (and should implement the same algorithm as ) the R function in DateUtils.R
' -----------------------------------------------------------------------------------------------------------------------
Function TenorToTime(ByVal Tenor As Variant)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res()

1         On Error GoTo ErrHandler
2         Force2DArrayR Tenor
3         NR = sNRows(Tenor): NC = sNCols(Tenor)

4         ReDim Res(1 To NR, 1 To NC)

5         For i = 1 To NR
6             For j = 1 To NC
7                 Res(i, j) = TenorToTimeCore(Tenor(i, j))
8             Next j
9         Next i
10        TenorToTime = Res

11        Exit Function
ErrHandler:
12        Throw "#TenorToTime (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Numeraire
' Author    : Philip Swannell
' Date      : 12-Oct-2016
' Purpose   : Returns the numeraire currency, taken from the market data workbook
' -----------------------------------------------------------------------------------------------------------------------
Function NumeraireFromMDWB()
1         On Error GoTo ErrHandler
2         NumeraireFromMDWB = RangeFromMarketDataBook("Config", "Numeraire")
3         Exit Function
ErrHandler:
4         Throw "#NumeraireFromMDWB (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FeedRatesFromTextFile
' Author     : Philip Swannell
' Date       : 21-Mar-2022
' Purpose    : Wrapper to method of the same name in the Market Data Workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub FeedRatesFromTextFile()
          
          Dim Ccys As String
          Dim FileName As String
          Dim MarketWB As Workbook
          Dim MethodName As String
          Dim origWindow As Window
          Dim SUH As clsScreenUpdateHandler
          Dim VersionNumber As Long

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()

3         Set MarketWB = OpenMarketWorkbook(True, False)
4         Set origWindow = ActiveWindow

5         VersionNumber = RangeFromSheet(MarketWB.Worksheets("Audit"), "Headers").Cells(2, 1).Value

6         If VersionNumber < gMinimumMarketDataWorkbookVersion Then
7             Throw "The market data workbook at '" & MarketWB.FullName & "' is an out-of-date version. It needs to be version " & CStr(gMinimumMarketDataWorkbookVersion) & " or later , but it is version " & VersionNumber
8         End If

9         FileName = RangeFromSheet(MarketWB.Worksheets("Config"), "MarketDataFile").Value
10        FileName = sJoinPath(MarketWB.Path, FileName)
11        MethodName = "'" & MarketWB.FullName & "'!FeedRatesFromTextFile"

12        Ccys = RangeFromSheet(shConfig, "CurrenciesToInclude")

13        ThrowIfError Application.Run(MethodName, FileName, "Cayley" & Ccys, False)

14        If Not ActiveWindow Is origWindow Then
15            origWindow.Activate
16        End If
17        Exit Sub
ErrHandler:
18        SomethingWentWrong "#FeedRatesFromTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub


