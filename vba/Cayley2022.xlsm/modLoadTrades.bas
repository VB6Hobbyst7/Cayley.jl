Attribute VB_Name = "modLoadTrades"
Option Explicit

Sub Test_LoadTradesFromTextFiles()

1         On Error GoTo ErrHandler
2         Application.ScreenUpdating = False
3         tic
4         LoadTradesFromTextFiles
5         toc "LoadTradesFromTextFiles"
6         Exit Sub
ErrHandler:
7         SomethingWentWrong "#Test_LoadTradesFromTextFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LoadTradesFromTextFiles
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Reads the contents of the three csv files provided by Airbus into sheets of the "trades workbook", which
'              use as an in-momory cache of trade data. The header names are morphed according to the three tables on
'              the StaticData sheet of this workbook and column order is changed according to preferred orders also
'              saved on that sheet. The tables mark certain column headers as required and errors are thrown if those
'              column headers are absent. The workbook into which the data is pasted is the CayleyTradesTemplate.xlsm
'              workbook which must exist on disk in the same folder as this workbook. That template file includes VBA code
'              for data sorting and filtering. After pasting data to the template file it's saved to disk as CayleyTrades.xlsm
'              in the same folder as this workbook.
' Parameters :
'  FxFile          : Name with path of the csv file containing the FX trades.
'  RatesFile       : Name with path of the csv file containing the Rates trades.
'  AmortisationFile: Name with path of the csv file containing amortisation details for rates trades.
'  HideOnOpening   : Once the data has been pasted to the workbook, should the window be hidden?
' -----------------------------------------------------------------------------------------------------------------------
Function LoadTradesFromTextFiles(Optional ByVal FxFile As String, Optional ByVal RatesFile As String, _
          Optional ByVal AmortisationFile As String, Optional HideOnOpening As Boolean) As Workbook

          Dim AmortisationContents
          Dim c As Range
          Dim FxContents
          Dim ODA
          Dim RatesContents
          Dim wb As Workbook
          Dim WorkbookFullName As String
          Dim ws As Worksheet
          Const TemplateFileName = "CayleyTradesTemplate.xlsm"
          Dim TemplateBookName As String
          Dim XSH As clsExcelStateHandler

1         On Error GoTo ErrHandler

2         Set XSH = CreateExcelStateHandler(, , False) 'disable events!

3         If FxFile = "" Then
4             FxFile = FileFromConfig("FxTradesCSVFile")
5         End If
6         If RatesFile = "" Then
7             RatesFile = FileFromConfig("RatesTradesCSVFile")
8         End If
9         If AmortisationFile = "" Then
10            AmortisationFile = FileFromConfig("AmortisationCSVFile")
11        End If

12        If Not sFileExists(FxFile) Then Throw "FxFile does not exist, we looked for it at '" + FxFile + "'"
13        If Not sFileExists(RatesFile) Then Throw "ratesFile does not exist, we looked for it at '" + RatesFile + "'"
14        If Not sFileExists(AmortisationFile) Then Throw "ratesFile does not exist, we looked for it at '" + AmortisationFile + "'"

15        StatusBarWrap "Parsing " + FxFile
16        FxContents = ThrowIfError(sCSVRead(FxFile, True, , , "d/m/y"))
17        FxContents = ConvertAndCheckHeaders(FxContents, RangeFromSheet(shStaticData, "FxFileHeaderConversion").Value, FxFile)

18        StatusBarWrap "Parsing " + RatesFile
19        RatesContents = ThrowIfError(sCSVRead(RatesFile, True, , , "d/m/y"))
20        RatesContents = ConvertAndCheckHeaders(RatesContents, RangeFromSheet(shStaticData, "RatesFileHeaderConversion").Value, RatesFile)

21        StatusBarWrap "Parsing " + AmortisationFile
22        AmortisationContents = ThrowIfError(sCSVRead(AmortisationFile, True, , , "d/m/y"))
23        AmortisationContents = ConvertAndCheckHeaders(AmortisationContents, RangeFromSheet(shStaticData, "AmortisationFileHeaderConversion").Value, AmortisationFile)

24        If IsInCollection(Application.Workbooks, gCayleyTradesWorkbookName) Then
25            Application.Workbooks(gCayleyTradesWorkbookName).Close False
26        End If
27        If IsInCollection(Application.Workbooks, TemplateFileName) Then
28            Application.Workbooks(TemplateFileName).Close False
29        End If

30        TemplateBookName = sJoinPath(ThisWorkbook.Path, TemplateFileName)
31        If Not sFileExists(TemplateBookName) Then Throw "Cannot find file '" + TemplateBookName + "'"

32        ODA = Application.DisplayAlerts
33        Application.DisplayAlerts = False

34        If Not sFileExists(TemplateBookName) Then
35            Throw "Sorry, we couldn't find '" & TemplateBookName & "'" & vbLf & _
                  "That's a workbook that should always be located in the same folder as " + ThisWorkbook.Name + "." & vbLf & _
                  "Is it possible it was moved, renamed or deleted?", True
36        End If

37        Set wb = Application.Workbooks.Open(TemplateBookName)
          'The template workbook should have only one worksheet, but just in case it was saved with more than one sheet, we delete
38        For Each ws In wb.Worksheets
39            If ws.Name <> "Audit" Then
40                ws.Delete
41            End If
42        Next

43        WorkbookFullName = sJoinPath(LocalTemp(), gCayleyTradesWorkbookName)

44        wb.SaveAs WorkbookFullName, xlOpenXMLWorkbookMacroEnabled
45        Application.DisplayAlerts = ODA

46        wb.Worksheets.Add wb.Worksheets("Audit"), , 4

47        LoadFromOneTradesFile wb.Worksheets(1), SN_FxTrades2, "Fx Trades", FxFile, FxContents, True, _
              RangeFromSheet(shStaticData, "FxPreferredOrder").Value, _
              RangeFromSheet(shStaticData, "FxFileHeaderConversion"), True
48        LoadFromOneTradesFile wb.Worksheets(2), SN_RatesTrades2, "Rates Trades", RatesFile, RatesContents, True, _
              RangeFromSheet(shStaticData, "RatesPreferredOrder").Value, _
              RangeFromSheet(shStaticData, "RatesFileHeaderConversion"), True

49        LoadFromOneTradesFile wb.Worksheets(3), SN_Amortisation2, "Amortisation", AmortisationFile, AmortisationContents, False

          Dim DataToPaste
          Dim Target As Range
50        Set ws = wb.Worksheets(4)
51        ws.Name = "DataSources"

52        With ws.Cells(1, 1)
53            .Value = gCayleyTradesWorkbookName
54            .Font.Size = 22
55        End With
          
56        DataToPaste = FileInfoWrapper(FxFile, "FxFile")
57        Set Target = ws.Cells(3, 1).Resize(sNRows(DataToPaste), sNCols(DataToPaste))
58        Target.Value = sArrayExcelString(DataToPaste)

59        DataToPaste = FileInfoWrapper(RatesFile, "RatesFile")
60        Set Target = Target.offset(Target.Rows.Count + 1).Resize(sNRows(DataToPaste), sNCols(DataToPaste))
61        Target.Value = sArrayExcelString(DataToPaste)

62        DataToPaste = FileInfoWrapper(AmortisationFile, "AmortisationFile")
63        Set Target = Target.offset(Target.Rows.Count + 1).Resize(sNRows(DataToPaste), sNCols(DataToPaste))
64        Target.Value = sArrayExcelString(DataToPaste)

65        For Each c In ws.UsedRange.Columns(1).Cells
66            If Not IsEmpty(c.offset(, 1).Value) Then
67                ws.Names.Add c.Value, c.offset(, 1)
68                If InStr(c.Value, "Date") > 0 Then
69                    c.offset(, 1).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
70                ElseIf InStr(c.Value, "Size") > 0 Then
71                    c.offset(, 1).NumberFormat = "#,##0;[Red]-#,##0"
72                End If
73            End If
74        Next c

75        With ws.UsedRange
76            .offset(1).Columns.AutoFit
77            .Columns(2).HorizontalAlignment = xlHAlignLeft
78        End With

79        ws.Activate
80        wb.Windows(1).DisplayHeadings = False
81        wb.Windows(1).DisplayGridlines = False

82        ws.Protect , True, True
            
83        wb.Worksheets(1).Activate
84        wb.Save

85        If HideOnOpening Then
86            wb.Windows(1).Visible = False
87        End If

88        Set XSH = Nothing 'Switches events back on
89        For Each ws In wb.Worksheets
              'Trigger the change event to update the "message cell"
90            If IsInCollection(ws.Names, "TheFilters") Then
91                ws.Range("TheFilters").ClearContents
92            End If
93        Next

94        StatusBarWrap False
            
95        wb.Saved = True

96        Set LoadTradesFromTextFiles = wb

97        Exit Function
ErrHandler:
98        Throw "#LoadTradesFromTextFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LoadFromOneTradesFile
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Sub of LoadTradesFromTextFiles. Handles one file to one worksheet.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub LoadFromOneTradesFile(ws As Worksheet, SheetName As String, Title As String, SourceFileName As String, _
          DataWithHeaders As Variant, WithSortButtons As Boolean, Optional PreferredColumnOrder As Variant, _
          Optional ConversionTable, Optional doFillIn As Boolean)

          Dim CommentText As String
          Dim i As Long
          Dim lo As ListObject
          Dim NumberFormat As String
          Dim OrigHeaders

1         On Error GoTo ErrHandler
2         ws.Name = SheetName
3         StatusBarWrap "Pasting to worksheet " + ws.Name

4         With ws.Cells(1, 1)
5             .Value = Title
6             .Font.Size = 22
7         End With
8         ws.Cells(2, 1).Value = "Source:"
9         ws.Cells(2, 2).Value = "'" & SourceFileName
10        ws.Cells(2, 1).Resize(1, 2).Font.ColorIndex = 16 'light grey

11        With ws.Cells(IIf(WithSortButtons, 7, 6), 1).Resize(sNRows(DataWithHeaders), sNCols(DataWithHeaders))
12            Application.GoTo .Cells(2, 1)
13            ActiveWindow.FreezePanes = True

14            .Value = DataWithHeaders

15            ws.Names.Add "TheDataWithHeaders", .Cells
16            If .Rows.Count > 1 Then
17                ws.Names.Add "TheData", .offset(1).Resize(.Rows.Count - 1)
18            End If

19            With .Rows(IIf(WithSortButtons, -1, 0))
20                AddGreyBorders .Cells
21                ws.Names.Add "TheFilters", .Cells
22                .NumberFormat = "@"
23                CayleyFormatAsInput .Cells
24            End With

25            If doFillIn Then
26                FillInMissingCounterpartyParents .Cells
27            End If
28            If Not IsMissing(ConversionTable) Then
29                OrigHeaders = sVLookup(.Rows(1).Value, ConversionTable, 1, 2)
30                For i = 1 To .Columns.Count
31                    If Not sIsErrorString(OrigHeaders(1, i)) Then
32                        If .Cells(1, i).Value <> OrigHeaders(1, i) Then
33                            CommentText = "In the source, this column is headed:" + vbLf + _
                                  "'" + OrigHeaders(1, i) + "'"
34                            SetCellComment .Cells(1, i), CommentText, False
35                        End If
36                    End If
37                Next i
38            End If

39            If .Rows.Count > 1 Then
40                For i = 1 To .Columns.Count
41                    NumberFormat = NumberFormatFromColumnHeader(.Cells(1, i).Value)
42                    If NumberFormat <> "General" Then
43                        .Columns(i).offset(1).Resize(.Rows.Count - 1).NumberFormat = NumberFormat
44                    End If
45                Next
46            End If

47            If Not IsMissing(PreferredColumnOrder) Then
48                SortColumns .Cells, PreferredColumnOrder
49            End If

50            Set lo = ws.ListObjects.Add(xlSrcRange, .Cells, , xlYes)
51            lo.ShowAutoFilterDropDown = False
52            .Columns.AutoFit
53            If WithSortButtons Then AddSortButtons .Rows(0)

54        End With
55        ws.Activate
56        ws.Parent.Windows(1).DisplayHeadings = False
57        ws.Parent.Windows(1).DisplayGridlines = False
58        ws.Protect , True, True, , , , , , , , , , , True, True

59        Exit Sub
ErrHandler:
60        Throw "#LoadFromOneTradesFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function NumberFormatFromColumnHeader(ByVal Header As String)

1         On Error GoTo ErrHandler
2         Header = LCase(Header)

3         If InStr(Header, "date") > 0 Then
4             NumberFormatFromColumnHeader = "dd-mmm-yyyy"
5         ElseIf InStr(Header, "principal") > 0 Or InStr(Header, "accrual") > 0 Or _
              InStr(Header, "notional") > 0 Or InStr(Header, "amount") > 0 Or _
              InStr(Header, "npv") > 0 Or Right(Header, 4) = " amt" Or InStr(Header, "clientpricer") > 0 Then
6             NumberFormatFromColumnHeader = "#,##0;[Red]-#,##0"
7         Else
8             NumberFormatFromColumnHeader = "General"
9         End If

10        Exit Function
ErrHandler:
11        Throw "#NumberFormatFromColumnHeader (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConvertAndCheckHeaders
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    : Morph the headers according to the ConversionTable, and throw an error is required headers are missing
' Parameters :
'  DataWithHeaders: Data already read from the file, 2d array
'  ConversionTable: Three columns: Header in file, header on sheet, is header required?
'  FileName       : Only used for error message generation.
' -----------------------------------------------------------------------------------------------------------------------
Function ConvertAndCheckHeaders(ByVal DataWithHeaders, ConversionTable, FileName As String)

          Dim ErrString As String
          Dim ExistingHeaders
          Dim i As Long
          Dim MissingHeaders
          Dim Plural As Boolean
          Dim RequiredHeaders
          Dim TranslatedHeaders

1         On Error GoTo ErrHandler

2         Force2DArrayR DataWithHeaders

3         If sNCols(ConversionTable) <> 3 Then Throw "ConversionTable must have three columns"

4         ExistingHeaders = sSubArray(DataWithHeaders, 1, 1, 1)

5         RequiredHeaders = sMChoose(sSubArray(ConversionTable, 1, 1, , 1), sArrayEquals(sSubArray(ConversionTable, 1, 3, , 1), True))

6         MissingHeaders = sCompareTwoArrays(sArrayTranspose(ExistingHeaders), RequiredHeaders, "21,CaseSensitive")
          
7         If sNRows(MissingHeaders) > 1 Then
8             MissingHeaders = sSubArray(MissingHeaders, 2)
9             Plural = sNRows(MissingHeaders) > 1
10            ErrString = sConcatenateStrings(sArrayConcatenate("'", MissingHeaders, "'"), ", ")
11            ErrString = "Header" & IIf(Plural, "s ", " ") & ErrString & IIf(Plural, " are", " is") & " required but not found in the top line of file " & FileName
12            Throw ErrString
13        End If

14        TranslatedHeaders = ThrowIfError(sVLookup(ExistingHeaders, ConversionTable))

15        For i = 1 To sNCols(TranslatedHeaders)
16            If sIsErrorString(TranslatedHeaders(1, i)) Then
17                TranslatedHeaders(1, i) = ExistingHeaders(1, i)
18            End If
19        Next i

20        For i = 1 To sNCols(TranslatedHeaders)
21            DataWithHeaders(1, i) = TranslatedHeaders(1, i)
22        Next

23        ConvertAndCheckHeaders = DataWithHeaders

24        Exit Function
ErrHandler:
25        Throw "#ConvertAndCheckHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileInfoWrapper
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    :
' Parameters :
'  FileName    :
'  FriendlyName:
' -----------------------------------------------------------------------------------------------------------------------
Function FileInfoWrapper(FileName As String, FriendlyName As String)
          Dim i As Long
          Dim Result

1         On Error GoTo ErrHandler
2         Result = sArrayStack(sArrayRange(FriendlyName, FileName), sFileInfo(FileName))
3         For i = 2 To sNRows(Result)
4             Result(i, 1) = FriendlyName & "_" & Result(i, 1)
5         Next i

6         FileInfoWrapper = Result

7         Exit Function
ErrHandler:
8         Throw "#FileInfoWrapper (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SortColumns
' Author     : Philip Swannell
' Date       : 04-Mar-2022
' Purpose    : Sort a range left to right, according to how the cells in row 1 of the range appear in the argument
'              PreferredOrder, which should be a 2d, 1-col array. Fails if cells just above RangeToSort are not blank
' -----------------------------------------------------------------------------------------------------------------------
Sub SortColumns(RangeToSort As Range, PreferredOrder As Variant)

          Dim c As Range
          Dim MatchRes
          Dim RangeWithExtraRow As Range

1         On Error GoTo ErrHandler
2         If RangeToSort.Row = 1 Then
3             Throw "The top of RangeToSort cannot be in row 1"
4         End If

5         For Each c In RangeToSort.Rows(0).Cells
6             If Not IsEmpty(c.Value) Then
7                 Throw "Cell immediately above RangeToSort must be blank"
8             End If
9         Next

10        MatchRes = sMatch(RangeToSort.Rows(1).Value, PreferredOrder)

11        RangeToSort.Rows(0).Value = MatchRes

12        Set RangeWithExtraRow = RangeToSort.offset(-1).Resize(RangeToSort.Rows.Count + 1)

13        RangeWithExtraRow.Sort RangeWithExtraRow.Rows(1), xlAscending, , , , , , xlNo, , , xlLeftToRight, xlPinYin, xlSortNormal

14        RangeWithExtraRow.Rows(1).ClearContents

15        Exit Sub
ErrHandler:
16        Throw "#SortColumns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FillInMissingCounterpartyParents
' Author     : Philip Swannell
' Date       : 11-Mar-2022
' Purpose    : In the trade files (csv) the important column Counterparty Parent is blank when the Counterparty is
'              "already a parent". But that's inconvenient. So this method patches up the contents of the Counterparty
'              Parent column (where blank) with the contents of the Counterparty column. Operation is done on the data
'              in the worksheet, operating on the cells of the worksheet.
' Parameters :
'  RWH:       The data in the sheet, including headers (think "Range With Headers")
' -----------------------------------------------------------------------------------------------------------------------
Sub FillInMissingCounterpartyParents(RWH As Range)
          Const CopyFromHeader As String = "Counterparty"
          Const CopyToHeader As String = "Counterparty Parent"
          Dim AnyFilledIn As Boolean
          Dim c As Range
          Dim ColOffset As Long
          Dim CopyFromCol As Long
          Dim CopyToCol As Long
          Dim HeadersT As Variant
          Dim MatchRes As Variant
          Dim pastethis As Variant
          
1         On Error GoTo ErrHandler
2         HeadersT = sArrayTranspose(RWH.Rows(1).Value)

3         MatchRes = sMatch(CopyFromHeader, HeadersT)
4         If IsNumber(MatchRes) Then
5             CopyFromCol = MatchRes
6             MatchRes = sMatch(CopyToHeader, HeadersT)
7             If IsNumber(MatchRes) Then
8                 CopyToCol = MatchRes
9                 ColOffset = CopyFromCol - CopyToCol

10                For Each c In RWH.Columns(CopyToCol).Cells
11                    If IsEmpty(c.Value) Then
12                        pastethis = c.offset(, ColOffset).Value
13                        If Not IsEmpty(pastethis) Then
14                            If VarType(pastethis) = vbString Then
15                                c.Value = "'" & pastethis
16                            Else
17                                c.Value = pastethis
18                            End If
19                            c.Font.ColorIndex = 3 'bright red
20                            AnyFilledIn = True
21                        End If
22                    End If
23                Next c
24                If AnyFilledIn Then
25                    With RWH.Cells(-3, 1)
26                        .Value = "Values for '" + CopyToHeader + "' shown in RED are blank in the source file," & _
                              " so the value for '" + CopyFromHeader + "' is substituted."
27                        .Font.ColorIndex = 16 'light grey
28                        .Characters(Start:=InStr(.Value, "RED"), Length:=3).Font.ColorIndex = 3
29                    End With
30                End If
31            End If
32        End If

33        Exit Sub
ErrHandler:
34        Throw "#FillInMissingCounterpartyParents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

