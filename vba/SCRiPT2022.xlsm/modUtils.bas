Attribute VB_Name = "modUtils"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : CounterpartiesFromJulia
' Author    : Philip Swannell
' Date      : 20-Apr-2016
' Purpose   : Returns a list of counterparties for which calculations have been done in Julia
'---------------------------------------------------------------------------------------
Function CounterpartiesFromJulia()
          Dim Allowed As Variant

1         On Error GoTo ErrHandler

2         Allowed = "#Not available!"
3         On Error Resume Next
4         Allowed = sSortedArray(sArrayTranspose(gResults("PartyExposures").keys))
5         On Error GoTo ErrHandler

6         If VarType(Allowed) = vbString Then
7             If Left(Allowed, 1) = "#" Then
8                 Throw "Unable to get list of Counterparties for which PFEs are available." + vbLf + "Have PFEs been calculated yet? ", True
9             Else
10                Force2DArray Allowed
11            End If
12        End If

13        CounterpartiesFromJulia = Allowed
14        Exit Function
ErrHandler:
15        Throw "#CounterpartiesFromJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CounterpartiesFromMarketBook
' Author    : Philip Swannell
' Date      : 21-Apr-2016
' Purpose   : Returns a list of counterparties on the CDS sheet of the market workbook, does not include "SELF" and may or maynot include WHATIF
'---------------------------------------------------------------------------------------
Function CounterpartiesFromMarketBook(IncludeWhatIf As Boolean)
          Dim Res
          Dim ResNS
          Dim shCDS As Worksheet

1         On Error GoTo ErrHandler
2         Set shCDS = OpenMarketWorkbook(True, False).Worksheets("Credit")
3         Res = sExpandDown(RangeFromSheet(shCDS, "CDSTopLeft").Offset(1, 0)).Value
4         ResNS = sMChoose(Res, sArrayNot(sArrayEquals(Res, gSELF)))
5         If IncludeWhatIf Then
6             CounterpartiesFromMarketBook = sArrayStack(ResNS, gWHATIF)
7         Else
8             CounterpartiesFromMarketBook = ResNS
9         End If

10        Exit Function
ErrHandler:
11        Throw "#CounterpartiesFromMarketBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' ----------------------------------------------------------------
' Procedure Name: IsArrayOfStrings
' Purpose: Tests if an arbitrary input is an array all of whose elements are strings
' Procedure Kind: Function
' Procedure Access: Private
' Parameter x ():
' Return Type: Boolean
' Author: Philip Swannell
' Date: 07-March-2018
' ----------------------------------------------------------------
Private Function IsArrayOfStrings(x) As Boolean
          Dim y As Variant
1         On Error GoTo ErrHandler
2         If IsArray(x) Then
3             IsArrayOfStrings = True
4             For Each y In x
5                 If VarType(y) <> vbString Then
6                     IsArrayOfStrings = False
7                     Exit Function
8                 End If
9             Next
10        End If
11        Exit Function
ErrHandler:
12        Throw "#IsArrayOfStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ChooseBanks
' Author    : Philip Swannell
' Date      : 29-Dec-2016
' Purpose   : Make it easy for the user to restrict the number of banks for which we do calculations...
'---------------------------------------------------------------------------------------
Function ChooseBanks(DoNothingIfChoiceAlreadyMade As Boolean, Optional ByRef BanksChosen) As Boolean
          Dim BanksProcessedLastRun
          Dim Res
          Dim TopText As String
          
1         On Error GoTo ErrHandler
          Static BanksLastChosen
          Dim Prepopulate
2         On Error Resume Next
3         BanksProcessedLastRun = sArrayTranspose(gResults("PartyResults")("PartyName"))
4         On Error GoTo ErrHandler

5         If Not IsArrayOfStrings(BanksProcessedLastRun) Then
6             BanksProcessedLastRun = CounterpartiesOnPortfolioSheet()
7         End If
8         Force2DArray BanksProcessedLastRun
9         If sNCols(BanksProcessedLastRun) > 1 Then
10            BanksProcessedLastRun = sArrayTranspose(BanksProcessedLastRun)
11        End If
12        If IsEmpty(BanksLastChosen) Then BanksLastChosen = BanksProcessedLastRun

13        If DoNothingIfChoiceAlreadyMade Then
14            If IsArrayOfStrings(BanksLastChosen) Then
15                BanksChosen = BanksLastChosen
16                ChooseBanks = True
17                Exit Function
18            End If
19        End If

          Const MiddleCaption = "All with trades" + vbCrLf + "and lines data"
          Dim ButtonClicked As String

20        TopText = "For which banks should XVAs be calculated" + vbLf + "and displayed on the xVADashboard?" + vbLf + vbLf + _
              "Choose fewer banks for faster run times."
21        Prepopulate = BanksLastChosen
TryAgain:

22        Res = ShowMultipleChoiceDialog(sSortedArray(CounterpartiesFromMarketBook(False)), Prepopulate, "Choose Banks", TopText, , , "OK", , Array(False, True), MiddleCaption, ButtonClicked)
23        If sIsErrorString(Res) Then
24            ChooseBanks = False
25            BanksChosen = Empty
26            Exit Function
27        Else
28            If ButtonClicked = MiddleCaption Then
29                Prepopulate = CounterpartiesOnPortfolioSheet()
30                GoTo TryAgain
31            End If
32            BanksChosen = Res
33            BanksLastChosen = Res
34        End If
35        ChooseBanks = True
36        Exit Function
ErrHandler:
37        Throw "#ChooseBanks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CounterpartiesOnPortfolioSheet
' Author    : Philip Swannell
' Date      : 10-Jan-2017
' Purpose   : Returns counterparties on portfolio sheet, sorted with duplicates removed, and WHATIF removed
'---------------------------------------------------------------------------------------
Function CounterpartiesOnPortfolioSheet()
          Dim Counterparties As Variant
          Dim NumTrades As Long
          Dim TR As Range

1         On Error GoTo ErrHandler
2         Set TR = getTradesRange(NumTrades)
3         If NumTrades = 0 Then
4             CounterpartiesOnPortfolioSheet = CreateMissing()
5             Exit Function
6         End If
7         Counterparties = TR.Columns(gCN_Counterparty).Value
8         Counterparties = sRemoveDuplicates(Counterparties, True)
9         If IsNumber(sMatch(gWHATIF, Counterparties)) Then
10            If sNRows(Counterparties) = 1 Then
11                CounterpartiesOnPortfolioSheet = CreateMissing()
12                Exit Function
13            Else
14                Counterparties = sMChoose(Counterparties, sArrayNot(sArrayEquals(Counterparties, gWHATIF)))
15            End If
16        End If

17        CounterpartiesOnPortfolioSheet = Counterparties

18        Exit Function
ErrHandler:
19        Throw "#CounterpartiesOnPortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromConfig
' Author     : Philip Swannell
' Date       : 14-Mar-2022
' Purpose    : For better relocatability, files and folders on the Config sheet may be given as paths relative to
'              the folder of this workbook. This function returns a file's full path.
' -----------------------------------------------------------------------------------------------------------------------
Function FileFromConfig(NameOnConfig As String)

          Dim Res As String
1         On Error GoTo ErrHandler
2         Res = RangeFromSheet(shConfig, NameOnConfig, False, True, False, False, False).Value
3         FileFromConfig = sJoinPath(ThisWorkbook.Path, Res)

4         Exit Function
ErrHandler:
5         Throw "#FileFromConfig (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

