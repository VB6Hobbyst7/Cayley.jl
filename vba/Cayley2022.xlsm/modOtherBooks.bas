Attribute VB_Name = "modOtherBooks"
Option Explicit
Option Base 1

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : NameForOpenOthers
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Name the menu option for opening more than one other workbook at a time.
' -----------------------------------------------------------------------------------------------------------------------
Function NameForOpenOthers(MarketBookIsOpen As Boolean, TradesBookIsOpen As Boolean, LinesBookIsOpen As Boolean, ForCreditUsageSheet As Boolean) As Variant
          Dim Flag As Long

1         On Error GoTo ErrHandler
2         Flag = IIf(TradesBookIsOpen, 0, 4) + IIf(MarketBookIsOpen, 0, 2) + IIf(LinesBookIsOpen, 0, 1)

3         Select Case Flag
              Case 7
4                 NameForOpenOthers = "&Open Trades, Market and Lines workbooks"
5             Case 6
6                 NameForOpenOthers = "&Open Trades and Market workbooks"
7             Case 5
8                 NameForOpenOthers = "&Open Trades and Lines workbooks"
9             Case 4
10                NameForOpenOthers = IIf(ForCreditUsageSheet, createmissing(), "&Open Trades workbook")
11            Case 3
12                NameForOpenOthers = "&Open Market and Lines workbooks"
13            Case 2
14                NameForOpenOthers = IIf(ForCreditUsageSheet, createmissing(), "&Open Market workbook")
15            Case 1
16                NameForOpenOthers = IIf(ForCreditUsageSheet, createmissing(), "&Open Lines workbook")
17            Case 0
18                NameForOpenOthers = createmissing()
19        End Select

20        Exit Function
ErrHandler:
21        Throw "#NameForOpenOthers (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OtherBooksAreOpen
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Returns TRUE if all of Market Data workbook, Trades workbook and Lines workbook are open
' -----------------------------------------------------------------------------------------------------------------------
Function OtherBooksAreOpen(ByRef MarketBookIsOpen As Boolean, ByRef TradesBookIsOpen, ByRef LinesBookIsOpen As Boolean) As Boolean
          Dim LinesBookName As String
          Dim MarketBookName As String

1         On Error GoTo ErrHandler

2         LinesBookName = FileFromConfig("LinesWorkbook")
3         MarketBookName = FileFromConfig("MarketDataWorkbook")

4         LinesBookName = ThrowIfError(sSplitPath(LinesBookName))
5         MarketBookName = ThrowIfError(sSplitPath(MarketBookName))

6         MarketBookIsOpen = IsInCollection(Application.Workbooks, MarketBookName)
7         TradesBookIsOpen = IsInCollection(Application.Workbooks, gCayleyTradesWorkbookName)
8         LinesBookIsOpen = IsInCollection(Application.Workbooks, LinesBookName)

9         OtherBooksAreOpen = MarketBookIsOpen And TradesBookIsOpen And LinesBookIsOpen

10        Exit Function
ErrHandler:
11        Throw "#OtherBooksAreOpen (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub OpenOtherBooks()
          Dim SUH As clsScreenUpdateHandler
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler
2         SetCalculationToManual
3         Set SUH = CreateScreenUpdateHandler()
4         Set XSH = CreateExcelStateHandler(, , False)

5         OpenLinesWorkbook True, False
6         OpenTradesWorkbook True, False
7         OpenMarketWorkbook True, False
8         Exit Sub
ErrHandler:
9         Throw "#OpenOtherBooks (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenMarketWorkbook
' Author    : Philip Swannell
' Date      : 26-Jul-2016
' Purpose   : Returns a workbook object representing the market data workbook whose
'             Fullname is given on the Config sheet of this workbook
' -----------------------------------------------------------------------------------------------------------------------
Function OpenMarketWorkbook(Optional HideOnOpening As Boolean, Optional Activate As Boolean) As Workbook
          Dim BookFullName As String
          Dim BookName As String
          Dim CopyOfErr As String
          Dim OldSaved As Boolean
          Dim origWn As Window
          Dim wb As Workbook
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler

          'This will prevent FlushStatics being called

2         BookFullName = FileFromConfig("MarketDataWorkbook")
3         BookName = sSplitPath(BookFullName)
4         Set origWn = ActiveWindow

5         If Not IsInCollection(Application.Workbooks, BookName) Then
6             If Not sFileExists(BookFullName) Then
7                 Throw "Sorry, we couldn't find '" & BookFullName & "'" & vbLf & _
                      "Is it possible it was moved, renamed or deleted?" & vbLf & _
                      "If so you need to change the ""MarketDataWorkbook"" " & _
                      "setting on the Config sheet of this workbook.", True
8             End If
9             If Not Activate Then Set XSH = CreateExcelStateHandler(, , False)
10            StatusBarWrap "Opening " & BookName
11            Set wb = Application.Workbooks.Open(BookFullName, , False, , , , , , , , , , True)
12            OldSaved = True
13            StatusBarWrap False

14            shConfig.Calculate        'since the Config sheet calls method RangeFromMarketDataBook
15            If HideOnOpening Then
16                wb.Windows(1).Visible = False
17            End If

18            If Not Activate Then
19                ThisWorkbook.Windows(1).Activate
20            End If
21        Else
22            Set wb = Application.Workbooks(BookName)
23            OldSaved = wb.Saved
24        End If

          'Check on compatibility
25        CheckMarketWorkbook wb, ThisWorkbook.Name, gMinimumMarketDataWorkbookVersion

26        If Activate Then
27            With wb.Windows(1)
28                .Visible = True
29                If .WindowState = xlMinimized Then .WindowState = xlNormal
30                .Activate
31            End With
32        Else
33            If Not ActiveWindow Is origWn Then origWn.Activate
34        End If

35        If OldSaved Then
36            If Not wb.Saved Then
37                wb.Saved = True
38            End If
39        End If

40        Set OpenMarketWorkbook = wb
41        Exit Function
ErrHandler:
42        CopyOfErr = "#OpenMarketWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
43        StatusBarWrap False
44        Throw CopyOfErr
End Function

Function OpenTradesWorkbook(HideOnOpening As Boolean, Activate As Boolean) As Workbook
          Dim BookFullName As String
          Dim BookName As String
          Dim CopyOfErr As String
          Dim origWn As Window
          Dim Res
          Dim wb As Workbook

1         On Error GoTo ErrHandler

2         BookName = gCayleyTradesWorkbookName
3         BookFullName = sJoinPath(LocalTemp(), BookName)
4         Set origWn = ActiveWindow

5         If IsInCollection(Application.Workbooks, BookName) Then
6             Set wb = Application.Workbooks(BookName)
7         Else
8             Set wb = LoadTradesFromTextFiles(, , , True)
                  
9             FlushStatics
10            Res = AllCurrenciesInTradesWorkBook(True)
11            If HideOnOpening Then
12                wb.Windows(1).Visible = False
13            End If
                        
14        End If

15        If Activate Then
16            With wb.Windows(1)
17                .Visible = True
18                If .WindowState = xlMinimized Then .WindowState = xlNormal
19                .Activate
20            End With
21        Else
22            If Not ActiveWindow Is origWn Then origWn.Activate
23        End If

24        Set OpenTradesWorkbook = wb
25        Exit Function
ErrHandler:
26        CopyOfErr = "#OpenTradesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
27        StatusBarWrap False
28        Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HasFileChanged
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Test if a file has changed since we previusly took a recod of its MD5, size and last modified date. Have
'              to guard against the possibility that caculating the MD5 hash fails, which happens if dot net
'              framework 3.5 is not installed and in this case we use only the file size and last modified date.
' -----------------------------------------------------------------------------------------------------------------------
Private Function HasFileChanged(FileName, OldMD5 As String, oldSize As Long, oldDateLastModified As Date) As Boolean
          Dim NewMD5 As String
          Dim useSizeDate As Boolean

1         On Error GoTo ErrHandler
2         If sIsErrorString(OldMD5) Then
3             useSizeDate = True
4         End If

5         If Not useSizeDate Then
6             NewMD5 = sFileCheckSum(FileName, "MD5")
7             If sIsErrorString(NewMD5) Then
8                 useSizeDate = True
9             End If
10        End If

11        If useSizeDate Then
12            HasFileChanged = (oldSize <> sFileInfo(FileName, "Size")) Or _
                  (oldDateLastModified <> sFileInfo(FileName, "DateLastModified"))
13        Else
14            HasFileChanged = NewMD5 <> OldMD5
15        End If

16        Exit Function
ErrHandler:
17        Throw "#HasFileChanged (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TradesWorkbookIsOutOfDate
' Author     : Philip Swannell
' Date       : 03-Mar-2022
' Purpose    : Returns true if we are in the "new world" of using CSV files for trades, the trades workbook is open and the
'              csv files have changed since they were last read into the trades workbook
' -----------------------------------------------------------------------------------------------------------------------
Function TradesWorkbookIsOutOfDate() As Boolean

          Dim AmortisationFile As String
          Dim FxFile As String
          Dim RatesFile As String
          Dim wb As Workbook
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         If IsInCollection(Application.Workbooks, gCayleyTradesWorkbookName) Then
            
3             FxFile = FileFromConfig("FxTradesCSVFile")
4             RatesFile = FileFromConfig("RatesTradesCSVFile")
5             AmortisationFile = FileFromConfig("AmortisationCSVFile")

6             If sFileExists(FxFile) Then
7                 If sFileExists(RatesFile) Then
8                     If sFileExists(AmortisationFile) Then
9                         Set wb = Application.Workbooks(gCayleyTradesWorkbookName)
10                        If IsInCollection(wb.Worksheets, "DataSources") Then
11                            Set ws = wb.Worksheets("DataSources")
12                            If HasFileChanged(RatesFile, RangeFromSheet(ws, "RatesFile_MD5").Value, _
                                  RangeFromSheet(ws, "RatesFile_Size").Value, _
                                  RangeFromSheet(ws, "RatesFile_DateLastModified").Value) Then
13                                TradesWorkbookIsOutOfDate = True
14                                Exit Function
15                            ElseIf HasFileChanged(FxFile, RangeFromSheet(ws, "FxFile_MD5").Value, _
                                  RangeFromSheet(ws, "FxFile_Size").Value, _
                                  RangeFromSheet(ws, "FxFile_DateLastModified").Value) Then
16                                TradesWorkbookIsOutOfDate = True
17                                Exit Function
18                            ElseIf HasFileChanged(AmortisationFile, RangeFromSheet(ws, "AmortisationFile_MD5").Value, _
                                  RangeFromSheet(ws, "AmortisationFile_Size").Value, _
                                  RangeFromSheet(ws, "AmortisationFile_DateLastModified").Value) Then
19                                TradesWorkbookIsOutOfDate = True
20                                Exit Function
21                            End If
22                        End If
23                    End If
24                End If
25            End If
26        End If

27        Exit Function
ErrHandler:
28        Throw "#TradesWorkbookIsOutOfDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AllCurrenciesInTradesWorkBook
' Author    : Philip Swannell
' Date      : 12-Sep-2016
' Purpose   : Returns column array of currencies of ALL trades in the trades workbook,
'             before any filtering is done.
' -----------------------------------------------------------------------------------------------------------------------
Function AllCurrenciesInTradesWorkBook(Optional ForceRefresh As Boolean = False)
          Static Res
1         On Error GoTo ErrHandler
2         If IsEmpty(Res) Or ForceRefresh Then
3             Res = CurrenciesFromQuery("None", "None", "None", "None", False, 0, True, True, _
                  OpenTradesWorkbook(True, False), shFutureTrades, Date)
4         End If
5         AllCurrenciesInTradesWorkBook = Res
6         Exit Function
ErrHandler:
7         Throw "#AllCurrenciesInTradesWorkBook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OpenLinesWorkbook
' Author    : Philip Swannell
' Date      : 28-Apr-2015
' Purpose   : Returns a workbook object representing the source workbook whose
'             Fullname is given on the main sheet of this workbook at the range "LinesWorkbook"
' -----------------------------------------------------------------------------------------------------------------------
Function OpenLinesWorkbook(Optional HideOnOpening As Boolean, Optional Activate As Boolean = False) As Workbook
          Dim BookFullName As String
          Dim BookName As String
          Dim CopyOfErr As String
          Dim HadToOpen As Boolean
          Dim origWn As Window
          Dim wb As Workbook
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler
            
2         BookFullName = FileFromConfig("LinesWorkbook")

3         BookName = sSplitPath(BookFullName)
4         Set origWn = ActiveWindow

5         On Error Resume Next
6         Set wb = Application.Workbooks(BookName)
7         On Error GoTo ErrHandler

8         If wb Is Nothing Then
9             If Not sFileExists(BookFullName) Then
10                Throw "Sorry, we couldn't find '" & BookFullName & "'" & vbLf & _
                      "Is it possible it was moved, renamed or deleted?" & vbLf _
                      & "If so you need to change the 'Lines Workbook'" & vbLf & _
                      "setting on the Config sheet of this workbook.", True
11            End If

12            If Not Activate Then Set XSH = CreateExcelStateHandler(, , False)
13            StatusBarWrap "Opening " & BookFullName
14            Set wb = Application.Workbooks.Open(BookFullName, , True, , , , , , , , , , True)
15            HadToOpen = True
16            StatusBarWrap False
              'Workbook will not open if "top of the call stack" is sheet calculation
17            If wb Is Nothing Then Throw "Cannot open " & BookFullName
18            CheckLinesWorkbook wb, True, True
19            SyncBanksInCayleyWithBanksInLinesBook
20            If HideOnOpening Then
21                wb.Windows(1).Visible = False
22            End If
23        End If

24        If Activate Then
25            With wb.Windows(1)
26                .Visible = True
27                If .WindowState = xlMinimized Then .WindowState = xlNormal
28                .Activate
29            End With
30        Else
31            If Not ActiveWindow Is origWn Then origWn.Activate
32        End If
33        If HadToOpen Then
34            wb.Saved = True
35        End If

36        Set OpenLinesWorkbook = wb
37        Exit Function
ErrHandler:
38        CopyOfErr = "#OpenLinesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
39        StatusBarWrap False
40        Throw CopyOfErr
End Function

Sub TestCheckLinesWorkbook()
1         On Error GoTo ErrHandler
2         CheckLinesWorkbook ActiveWorkbook, True, True
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestCheckLinesWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetColumnFromLinesBook
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Returns the contents of an entire column from the Lines book
' -----------------------------------------------------------------------------------------------------------------------
Function GetColumnFromLinesBook(ByVal Header As String, LinesBook As Workbook)
1         On Error GoTo ErrHandler

          Dim ColNumber As Variant
          Dim EntireRange As Range
          Dim EntireRangeNoHeaders As Range
          Dim HeaderRow As Range
          Dim HeaderRowTranspose

2         GetLinesRanges EntireRange, EntireRangeNoHeaders, HeaderRow, LinesBook
3         HeaderRowTranspose = Application.WorksheetFunction.Transpose(HeaderRow.Value2)
4         ColNumber = sMatch(Header, HeaderRowTranspose)
5         If Not IsNumber(ColNumber) Then Throw "Cannot find column headed '" & Header & "' in the Lines Workbook"
6         GetColumnFromLinesBook = EntireRangeNoHeaders.Columns(ColNumber).Value

7         Exit Function
ErrHandler:
8         Throw "#GetColumnFromLinesBook (line " & CStr(Erl) & "): " & Err.Description & "!"
9     End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetLinesRanges
' Author    : Philip Swannell
' Date      : 09-Nov-2016
' Purpose   : Subroutine shared between LookupCounterpartyInfo and GetColumnFromLinesBook
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLinesRanges(ByRef EntireRange As Range, ByRef EntireRangeNoHeaders As Range, _
          ByRef HeaderRange As Range, LinesBook As Workbook)
          Dim SheetName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         SheetName = SN_Lines

3         If Not IsInCollection(LinesBook.Worksheets, SheetName) Then
4             Throw "Lines workbook (" & LinesBook.Name & ") has no worksheet named " & SheetName
5         End If
6         Set ws = LinesBook.Worksheets(SheetName)
          Dim lo As ListObject

7         Select Case ws.ListObjects.Count
              Case 0
8                 Throw SN_Lines & " in workbook " & LinesBook.Name & _
                      " must have a 'Table' (Ribbon > Insert > Table) containing the lines data."
9             Case 1
10                Set lo = ws.ListObjects(1)
11                Set EntireRangeNoHeaders = lo.DataBodyRange
12                Set HeaderRange = lo.HeaderRowRange
13                Set EntireRange = Application.Union(HeaderRange, EntireRangeNoHeaders)
14            Case Else
15                Throw SN_Lines & " in workbook " & LinesBook.Name & _
                      " must have just one 'Table' (Ribbon > Insert > Table) but it has " & CStr(ws.ListObjects.Count)
16        End Select
17        Exit Function
ErrHandler:
18        Throw "#GetLinesRanges (line " & CStr(Erl) & "): " & Err.Description & "!"
19    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LookupCounterpartyInfo
' Author    : Philip Swannell
' Date      : 28-Apr-2015
' Purpose   : Get information from the Lines source book about a given bank or banks.
'             Both ParentCounterparty and HeaderNames can be passed as single column arrays.
'             When ParentCounterparty has more than one element, the return has more than one COLUMN
' -----------------------------------------------------------------------------------------------------------------------
Function LookupCounterpartyInfo(ByVal ParentCounterparty As Variant, ByVal HeaderNames As Variant, _
          Optional ReplaceEmptyWith As Variant = "#Empty data found!", Optional ReplaceBankNotFoundWith)

          Dim AllBanks As Variant
          Dim ColNumbers As Variant
          Dim EntireRange As Range
          Dim EntireRangeNoHeaders As Range
          Dim HeaderRange As Range
          Dim HeaderRowTranspose As Variant
          Dim i As Long
          Dim LinesBook As Workbook
          Dim MatchID
          Dim RowNumbers As Variant

1         On Error GoTo ErrHandler

2         Set LinesBook = OpenLinesWorkbook(True, False)

3         Force2DArrayRMulti HeaderNames, ParentCounterparty
4         If sNCols(HeaderNames) > 1 Then
5             Throw "HeaderNames must be entered as a single column array"
6         End If

7         GetLinesRanges EntireRange, EntireRangeNoHeaders, HeaderRange, LinesBook

8         HeaderRowTranspose = Application.WorksheetFunction.Transpose(HeaderRange.Value2)

9         MatchID = sMatch("CPTY_PARENT", HeaderRowTranspose)
10        If VarType(MatchID) = vbString Then Throw "Cannot find column headed ""CPTY_PARENT"" in the Lines workbook"

11        AllBanks = EntireRangeNoHeaders.Columns(MatchID).Value2
12        RowNumbers = sMatch(ParentCounterparty, AllBanks)
13        ColNumbers = sMatch(HeaderNames, HeaderRowTranspose)

14        Force2DArray ColNumbers
15        Force2DArray RowNumbers
          Dim j As Long
          Dim OutputNumCols As Long
          Dim OutputNumRows As Long
16        OutputNumCols = sNRows(RowNumbers)
17        OutputNumRows = sNRows(ColNumbers)

          Dim Result()
18        ReDim Result(1 To OutputNumRows, 1 To OutputNumCols)

19        For j = 1 To OutputNumCols
20            For i = 1 To OutputNumRows
21                If VarType(RowNumbers(j, 1)) = vbString Then
22                    If CStr(HeaderNames(i, 1)) = "Very short name" Then
23                        Result(i, j) = CStr(ParentCounterparty(j, 1))
24                    ElseIf Not IsMissing(ReplaceBankNotFoundWith) Then
25                        Result(i, j) = ReplaceBankNotFoundWith
26                    Else
27                        Throw "#Cannot find bank '" & CStr(ParentCounterparty(j, 1))
28                    End If
29                ElseIf VarType(ColNumbers(i, 1)) = vbString Then
30                    Throw "#Cannot find header '" & CStr(HeaderNames(i, 1)) & "' in lines workbook!"
31                Else
32                    Result(i, j) = EntireRangeNoHeaders.Cells(RowNumbers(j, 1), ColNumbers(i, 1)).Value2
33                    If IsEmpty(Result(i, j)) Then
34                        Result(i, j) = ReplaceEmptyWith
35                    End If
36                End If
37            Next i
38        Next j

39        LookupCounterpartyInfo = Result
40        Set LinesBook = Nothing
41        Exit Function
ErrHandler:
42        Throw "#LookupCounterpartyInfo (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PreferredCcy
' Author    : Philip Swannell
' Date      : 22-May-2015
' Purpose   : For line-utilisation based on trade notional we have to choose which of the two notionals to look at
'             This function encasulates the choice. Use the bank's base ccy if it's one of the currencies of the trade
'             otherwise look at a hard-coded order of preference
' -----------------------------------------------------------------------------------------------------------------------
Function PreferredCcy(ByVal CCY1 As String, ByVal CCY2 As String, ByVal BaseCCY As String)
          Static PreferenceArray As Variant
          Dim c As Variant

1         On Error GoTo ErrHandler
2         CCY1 = UCase(CCY1): CCY2 = UCase(CCY2): BaseCCY = UCase(BaseCCY)
3         If CCY1 = BaseCCY Or CCY2 = BaseCCY Then
4             PreferredCcy = BaseCCY
5             Exit Function
6         End If
7         If IsEmpty(PreferenceArray) Then
              'Order by weighting in Cayley's portfolio
8             PreferenceArray = Array("EUR", "USD", "GBP", "SAR", "BRL", "CAD", "AUD", "ZAR", "JPY", _
                  "NOK", "SEK", "CZK", "DKK", "RUB", "RON", "SGD", "CHF", "HUF", "PLN", "AED", "QAR")
9         End If
10        For Each c In PreferenceArray
11            If CCY1 = c Or CCY2 = c Then
12                PreferredCcy = c
13                Exit Function
14            End If
15        Next c
          'Not in list...
16        If CCY1 < CCY2 Then
17            PreferredCcy = CCY1
18        Else
19            PreferredCcy = CCY2
20        End If
21        Exit Function
ErrHandler:
22        Throw "#PreferredCcy (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub TestCheckMarketWorkbook()
1         On Error GoTo ErrHandler
2         CheckMarketWorkbook ActiveWorkbook, "Cayley"
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestCheckMarketWorkbook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


