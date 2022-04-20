Attribute VB_Name = "modCSVReadWrite"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme
' This version at: https://github.com/PGS62/VBA-CSV/releases/tag/v0.15

'CHANGES
'1) Comment out ThrowIfError
'2) CSVRead -> sCSVRead
'3) CSVWrite -> sCSVWrite
'4) Have added arguments TrueString and FalseString to CSVWrite

Option Explicit

Private m_FSO As Scripting.FileSystemObject

#If VBA7 And Win64 Then
'for 64-bit Excel
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
'for 32-bit Excel
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private Enum enmErrorStyle
    es_ReturnString = 0
    es_RaiseError = 1
End Enum

Private Const m_ErrorStyle As Long = es_ReturnString

Private Const m_LBound As Long = 1 'Sets the array lower bounds of the return from sCSVRead.
'To return zero-based arrays, change the value of this constant to 0.

Private Enum enmSourceType
    st_File = 0
    st_URL = 1
    st_String = 2
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path, or else a URL of a file, or else a string in
'             CSV format.
' ConvertTypes: Controls whether fields in the file are converted to typed values or remain as strings, and
'             sets the treatment of "quoted fields" and space characters.
'
'             ConvertTypes should be a string of zero or more letters from allowed characters `NDBETQ`.
'
'             The most commonly useful letters are:
'             1) `N` number fields are returned as numbers (Doubles).
'             2) `D` date fields (that respect DateFormat) are returned as Dates.
'             3) `B` fields matching TrueStrings or FalseStrings are returned as Booleans.
'
'             ConvertTypes is optional and defaults to the null string for no type conversion. `TRUE` is
'             equivalent to `NDB` and `FALSE` to the null string.
'
'             Three further options are available:
'             4) `E` fields that match Excel errors are converted to error values. There are fourteen of
'             these, including `#N/A`, `#NAME?`, `#VALUE!` and `#DIV/0!`.
'             5) `T` leading and trailing spaces are trimmed from fields. In the case of quoted fields,
'             this will not remove spaces between the quotes.
'             6) `Q` conversion happens for both quoted and unquoted fields; otherwise only unquoted fields
'             are converted.
'
'             For most files, correct type conversion can be achieved with ConvertTypes as a string which
'             applies for all columns, but type conversion can also be specified on a per-column basis.
'
'             Enter an array (or range) with two columns or two rows, column numbers on the left/top and
'             type conversion (subset of `NDBETQ`) on the right/bottom. Instead of column numbers, you can
'             enter strings matching the contents of the header row, and a column number of zero applies to
'             all columns not otherwise referenced.
'
'             For convenience when calling from VBA, you can pass an array of two element arrays such as
'             `Array(Array(0,"N"),Array(3,""),Array("Phone",""))` to convert all numbers in a file into
'             numbers in the return except for those in column 3 and in the column(s) headed "Phone".
' Delimiter : By default, sCSVRead will try to detect a file's delimiter as the first instance of comma, tab,
'             semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the
'             file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a
'             different character or string delimiter you should pass that as the Delimiter argument.
'
'             Alternatively, enter `FALSE` as the delimiter to treat the file as "not a delimited file". In
'             this case the return will mimic how the file would appear in a text editor such as NotePad.
'             The file will be split into lines at all line breaks (irrespective of double quotes) and each
'             element of the return will be a line of the file.
' IgnoreRepeated: Whether delimiters which appear at the start of a line, the end of a line or immediately
'             after another delimiter should be ignored while parsing; useful for fixed-width files with
'             delimiter padding between fields.
' DateFormat: The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO`
'             for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g.
'             2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time.
' Comment   : Rows that start with this string will be skipped while parsing.
' IgnoreEmptyLines: Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each
'             column will be assigned ShowMissingsAs for that empty row).
' HeaderRowNum: The row in the file containing headers. Type conversion is not applied to fields in the
'             header row, though leading and trailing spaces are trimmed.
'
'             This argument is most useful when calling from VBA, with SkipToRow set to one more than
'             HeaderRowNum. In that case the function returns the rows starting from SkipToRow, and the
'             header row is returned via the by-reference argument HeaderRow. Optional and defaults to 0.
' SkipToRow : The first row in the file that's included in the return. Optional and defaults to one more
'             than HeaderRowNum.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from the
'             first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the
'             end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol
'             are read.
' TrueStrings: Indicates how `TRUE` values are represented in the file. May be a string, an array of strings
'             or a range containing strings; by default, `TRUE`, `True` and `true` are recognised.
' FalseStrings: Indicates how `FALSE` values are represented in the file. May be a string, an array of
'             strings or a range containing strings; by default, `FALSE`, `False` and `false` are
'             recognised.
' MissingStrings: Indicates how missing values are represented in the file. May be a string, an array of
'             strings or a range containing strings. By default, only an empty field (consecutive
'             delimiters) is considered missing.
' ShowMissingsAs: Fields which are missing in the file (consecutive delimiters) or match one of the
'             MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null
'             string or `#N/A!` error value can be good alternatives.
'
'             If NumRows is greater than the number of rows in the file then the return is "padded" with
'             the value of ShowMissingsAs. Likewise, if NumCols is greater than the number of columns in
'             the file.
' Encoding  : Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be
'             omitted and sCSVRead will detect the file's encoding. If auto-detection does not work then
'             it's possible that the file is encoded `UTF-8` or `UTF-16` but without a byte option mark to
'             identify the encoding. Experiment with Encoding as each of `UTF-8` and `UTF-16`.
' DecimalSeparator: In many places in the world, floating point number decimals are separated with a comma
'             instead of a period (3,14 vs. 3.14). sCSVRead can correctly parse these numbers by passing in
'             the DecimalSeparator as a comma, in which case comma ceases to be a candidate if the parser
'             needs to guess the Delimiter.
' HeaderRow : This by-reference argument is for use from VBA (as opposed to from Excel formulas). It is
'             populated with the contents of the header row, with no type conversion, though leading and
'             trailing spaces are removed.
'
' Notes     : See also companion function sCSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Public Function sCSVRead(ByVal FileName As String, Optional ByVal ConvertTypes As Variant = False, _
        Optional ByVal Delimiter As Variant, Optional ByVal IgnoreRepeated As Boolean, _
        Optional ByVal DateFormat As String = "Y-M-D", Optional ByVal Comment As String, _
        Optional ByVal IgnoreEmptyLines As Boolean, Optional ByVal HeaderRowNum As Long, _
        Optional ByVal SkipToRow As Long, Optional ByVal SkipToCol As Long = 1, _
        Optional ByVal NumRows As Long, Optional ByVal NumCols As Long, _
        Optional ByVal TrueStrings As Variant, Optional ByVal FalseStrings As Variant, _
        Optional ByVal MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant, _
        Optional ByVal Encoding As Variant, Optional ByVal DecimalSeparator As String, _
        Optional ByRef HeaderRow As Variant) As Variant
Attribute sCSVRead.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute sCSVRead.VB_ProcData.VB_Invoke_Func = " \n26"

          Const DQ As String = """"
          Const Err_Delimiter As String = "Delimiter character must be passed as a string, FALSE for no delimiter. " & _
              "Omit to guess from file contents"
          Const Err_Delimiter2 As String = "Delimiter must have at least one character and cannot start with a double " & _
              "quote, line feed or carriage return"
          Const Err_FileEmpty As String = "File is empty"
          Const Err_FunctionWizard  As String = "Disabled in Function Wizard"
          Const Err_NumCols As String = "NumCols must be positive to read a given number of columns, or zero or omitted " & _
              "to read all columns from SkipToCol to the maximum column encountered."
          Const Err_NumRows As String = "NumRows must be positive to read a given number of rows, or zero or omitted to " & _
              "read all rows from SkipToRow to the end of the file."
          Const Err_Seps1 As String = "DecimalSeparator must be a single character"
          Const Err_Seps2 As String = "DecimalSeparator must not be equal to the first character of Delimiter or to a " & _
              "line-feed or carriage-return"
          Const Err_SkipToCol As String = "SkipToCol must be at least 1."
          Const Err_SkipToRow As String = "SkipToRow must be at least 1."
          Const Err_Comment As String = "Comment must not contain double-quote, line feed or carriage return"
          Const Err_HeaderRowNum As String = "HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow"
          
          Dim AcceptWithoutTimeZone As Boolean
          Dim AcceptWithTimeZone As Boolean
          Dim Adj As Long
          Dim AnyConversion As Boolean
          Dim AnySentinels As Boolean
          Dim CallingFromWorksheet As Boolean
          Dim CharSet As String
          Dim ColByColFormatting As Boolean
          Dim ColIndexes() As Long
          Dim ConvertQuoted As Boolean
          Dim CSVContents As String
          Dim CTDict As Scripting.Dictionary
          Dim DateOrder As Long
          Dim DateSeparator As String
          Dim ErrRet As String
          Dim Err_StringTooLong As String
          Dim i As Long
          Dim ISO8601 As Boolean
          Dim j As Long
          Dim k As Long
          Dim Lengths() As Long
          Dim M As Long
          Dim MaxSentinelLength As Long
          Dim MSLIA As Long
          Dim NeedToFill As Boolean
          Dim NotDelimited As Boolean
          Dim NumColsFound As Long
          Dim NumColsInReturn As Long
          Dim NumFields As Long
          Dim NumRowsFound As Long
          Dim NumRowsInReturn As Long
          Dim QuoteCounts() As Long
          Dim Ragged As Boolean
          Dim ReturnArray() As Variant
          Dim RowIndexes() As Long
          Dim Sentinels As Scripting.Dictionary
          Dim SepStandard As Boolean
          Dim ShowBooleansAsBooleans As Boolean
          Dim ShowDatesAsDates As Boolean
          Dim ShowErrorsAsErrors As Boolean
          Dim ShowMissingsAsEmpty As Boolean
          Dim ShowNumbersAsNumbers As Boolean
          Dim SourceType As enmSourceType
          Dim Starts() As Long
          Dim strDelimiter As String
          Dim Stream As Object 'either ADODB.Stream or Scripting.TextStram
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
          Dim SysDecimalSeparator As String
          Dim TempFile As String
          Dim TrimFields As Boolean
          Dim TriState As Long
          Dim useADODB As Boolean
          
1         On Error GoTo ErrHandler

2         SourceType = InferSourceType(FileName)

          'Download file from internet to local temp folder
3         If SourceType = st_URL Then
4             TempFile = Environ$("Temp") & "\VBA-CSV\Downloads\DownloadedFile.csv"
5             FileName = Download(FileName, TempFile)
6             SourceType = st_File
7         End If

          'Parse and validate inputs...
8         If SourceType <> st_String Then
9             ParseEncoding FileName, Encoding, TriState, CharSet, useADODB
10        End If

11        If VarType(Delimiter) = vbBoolean Then
12            If Not Delimiter Then
13                NotDelimited = True
14            Else
15                Throw Err_Delimiter
16            End If
17        ElseIf VarType(Delimiter) = vbString Then
18            If Len(Delimiter) = 0 Then
19                strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
20            ElseIf Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
21                Throw Err_Delimiter2
22            Else
23                strDelimiter = Delimiter
24            End If
25        ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
26            strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
27        Else
28            Throw Err_Delimiter
29        End If

30        SysDecimalSeparator = Application.DecimalSeparator
31        If DecimalSeparator = vbNullString Then DecimalSeparator = SysDecimalSeparator
32        If DecimalSeparator = SysDecimalSeparator Then
33            SepStandard = True
34        ElseIf Len(DecimalSeparator) <> 1 Then
35            Throw Err_Seps1
36        ElseIf DecimalSeparator = strDelimiter Or DecimalSeparator = vbLf Or DecimalSeparator = vbCr Then
37            Throw Err_Seps2
38        End If

39        Set CTDict = New Scripting.Dictionary

40        ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
              ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, _
              ConvertQuoted, TrimFields, ColByColFormatting, HeaderRowNum, CTDict

41        Set Sentinels = New Scripting.Dictionary
42        MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
              ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings
          
43        If ShowDatesAsDates Then
44            ParseDateFormat DateFormat, DateOrder, DateSeparator, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone
45            SysDateOrder = Application.International(xlDateOrder)
46            SysDateSeparator = Application.International(xlDateSeparator)
47        End If

48        If HeaderRowNum < 0 Then Throw Err_HeaderRowNum
49        If SkipToRow = 0 Then SkipToRow = HeaderRowNum + 1
50        If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
51        If SkipToCol = 0 Then SkipToCol = 1
52        If SkipToRow < 1 Then Throw Err_SkipToRow
53        If SkipToCol < 1 Then Throw Err_SkipToCol
54        If NumRows < 0 Then Throw Err_NumRows
55        If NumCols < 0 Then Throw Err_NumCols

56        If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
             
57        If InStr(Comment, DQ) > 0 Or InStr(Comment, vbLf) > 0 Or InStr(Comment, vbCrLf) > 0 Then Throw Err_Comment
          'End of input validation
          
58        CallingFromWorksheet = TypeName(Application.Caller) = "Range"
          
59        If CallingFromWorksheet Then
60            If SourceType = st_File Then
61                If FunctionWizardActive() Then
62                    sCSVRead = "#" & Err_FunctionWizard & "!"
63                    Exit Function
64                End If
65            End If
66        End If
          
67        If NotDelimited Then
68            sCSVRead = ParseTextFile(FileName, SourceType <> st_String, useADODB, CharSet, TriState, SkipToRow, NumRows, CallingFromWorksheet)
69            Exit Function
70        End If
                
71        If SourceType = st_String Then
72            CSVContents = FileName
              
73            ParseCSVContents CSVContents, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                  IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, _
                  NumFields, Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
74        Else
75            If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
                  
76            If useADODB Then
77                Set Stream = CreateObject("ADODB.Stream")
78                Stream.CharSet = CharSet
79                Stream.Open
80                Stream.LoadFromFile FileName
81                If Stream.EOS Then Throw Err_FileEmpty
82            Else
83                Set Stream = m_FSO.GetFile(FileName).OpenAsTextStream(ForReading, TriState)
84                If Stream.atEndOfStream Then Throw Err_FileEmpty
85            End If

86            If SkipToRow = 1 And NumRows = 0 Then
87                CSVContents = ReadAllFromStream(Stream)
88                Stream.Close
                  
89                ParseCSVContents CSVContents, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                      IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                      Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
90            Else
91                CSVContents = ParseCSVContents(Stream, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                      IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                      Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow)
92                Stream.Close
93            End If
94        End If
                 
          '    'Useful for debugging method ParseCSVContents - return an array displaying the variables set by that method
          '    Dim Chars() As String, Numbers() As Long, Ascs() As Long
          '    ReDim Numbers(1 To Len(CSVContents), 1 To 1)
          '    ReDim Chars(1 To Len(CSVContents), 1 To 1)
          '    ReDim Ascs(1 To Len(CSVContents), 1 To 1)
          '    For i = 1 To Len(CSVContents)
          '        Chars(i, 1) = Mid$(CSVContents, i, 1)
          '        Numbers(i, 1) = i
          '        Ascs(i, 1) = AscW(Chars(i, 1))
          '    Next i
          '    Dim Headers
          '    Headers = HStack("NRF,NCF,NF,Dlm", "Starts", "Lengths", "RowIndexes", "ColIndexes", "QuoteCounts", "i", "Char(i)", "AscW(Char(i))")
          '    sCSVRead = VStack(Headers, HStack(VStack(NumRowsFound, NumColsFound, NumFields, strDelimiter), Transpose(Starts), _
          '        Transpose(Lengths), Transpose(RowIndexes), Transpose(ColIndexes), Transpose(QuoteCounts), Numbers, Chars, Ascs))
          '    Exit Function
                 
95        If NumCols = 0 Then
96            NumColsInReturn = NumColsFound - SkipToCol + 1
97            If NumColsInReturn <= 0 Then
98                Throw "SkipToCol (" & CStr(SkipToCol) & _
                      ") exceeds the number of columns in the file (" & CStr(NumColsFound) & ")"
99            End If
100       Else
101           NumColsInReturn = NumCols
102       End If
103       If NumRows = 0 Then
104           NumRowsInReturn = NumRowsFound
105       Else
106           NumRowsInReturn = NumRows
107       End If
              
108       AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowBooleansAsBooleans Or ShowErrorsAsErrors Or TrimFields
              
109       Adj = m_LBound - 1
110       ReDim ReturnArray(1 + Adj To NumRowsInReturn + Adj, 1 + Adj To NumColsInReturn + Adj)
111       MSLIA = MaxStringLengthInArray()
112       ShowMissingsAsEmpty = IsEmpty(ShowMissingsAs)
              
113       For k = 1 To NumFields
114           i = RowIndexes(k)
115           j = ColIndexes(k) - SkipToCol + 1
116           If j >= 1 And j <= NumColsInReturn Then
117               If CallingFromWorksheet Then
118                   If Lengths(k) > MSLIA Then
119                       Err_StringTooLong = "The file has a field (row " + CStr(i + SkipToRow - 1) & _
                              ", column " & CStr(j + SkipToCol - 1) & ") of length " + Format$(Lengths(k), "###,###")
120                       If MSLIA >= 32767 Then
121                           Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " + Format$(MSLIA, "####,####")
122                       Else
123                           Err_StringTooLong = Err_StringTooLong & _
                                  ". An array containing a string longer than " + Format$(MSLIA, "###,###") + _
                                  " cannot be returned from VBA to an Excel worksheet"
124                       End If
125                       Throw Err_StringTooLong
126                   End If
127               End If
              
128               If ColByColFormatting Then
129                   ReturnArray(i + Adj, j + Adj) = Mid$(CSVContents, Starts(k), Lengths(k))
130               Else
131                   ReturnArray(i + Adj, j + Adj) = ConvertField(Mid$(CSVContents, Starts(k), Lengths(k)), AnyConversion, _
                          Lengths(k), TrimFields, DQ, QuoteCounts(k), ConvertQuoted, ShowNumbersAsNumbers, SepStandard, _
                          DecimalSeparator, SysDecimalSeparator, ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, _
                          AcceptWithTimeZone, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, _
                          Sentinels, MaxSentinelLength, ShowMissingsAs)
132               End If
                  
                  'File has variable number of fields per line...
133               If Ragged Then
134                   If Not ShowMissingsAsEmpty Then
135                       If k = NumFields Then
136                           NeedToFill = j < NumColsInReturn
137                       ElseIf RowIndexes(k + 1) > RowIndexes(k) Then
138                           NeedToFill = j < NumColsInReturn
139                       Else
140                           NeedToFill = False
141                       End If
142                       If NeedToFill Then
143                           For M = j + 1 To NumColsInReturn
144                               ReturnArray(i + Adj, M + Adj) = ShowMissingsAs
145                           Next M
146                       End If
147                   End If
148               End If
149           End If
150       Next k

151       If Ragged Then
152           If Not IsEmpty(HeaderRow) Then
153               If NCols(HeaderRow) < NCols(ReturnArray) + SkipToCol - 1 Then
154                   ReDim Preserve HeaderRow(1 To 1, 1 To NCols(ReturnArray) + SkipToCol - 1)
155               End If
156           End If
157       End If
158       If SkipToCol > 1 Then
159           If Not IsEmpty(HeaderRow) Then
                  Dim HeaderRowTruncated() As String
160               ReDim HeaderRowTruncated(1 To 1, 1 To NumColsInReturn)
161               For i = 1 To NumColsInReturn
162                   HeaderRowTruncated(1, i) = HeaderRow(1, i + SkipToCol - 1)
163               Next i
164               HeaderRow = HeaderRowTruncated
165           End If
166       End If
          
          'In this case no type conversion should be applied to the top row of the return
167       If HeaderRowNum = SkipToRow Then
168           If AnyConversion Then
169               For i = 1 To MinLngs(NCols(HeaderRow), NumColsInReturn)
170                   ReturnArray(1 + Adj, i + Adj) = HeaderRow(1, i)
171               Next
172           End If
173       End If

174       If ColByColFormatting Then
              Dim CT As Variant
              Dim Field As String
              Dim NC As Long
              Dim NCH As Long
              Dim NR As Long
              Dim QC As Long
              Dim UnQuotedHeader As String
175           NR = NRows(ReturnArray)
176           NC = NCols(ReturnArray)
177           If IsEmpty(HeaderRow) Then
178               NCH = 0
179           Else
180               NCH = NCols(HeaderRow) 'possible that headers has fewer than expected columns if file is ragged
181           End If

182           For j = 1 To NC
183               If j + SkipToCol - 1 <= NCH Then
184                   UnQuotedHeader = HeaderRow(1, j + SkipToCol - 1)
185               Else
186                   UnQuotedHeader = -1 'Guaranteed not to be a key of the Dictionary
187               End If
188               If CTDict.Exists(j + SkipToCol - 1) Then
189                   CT = CTDict.item(j + SkipToCol - 1)
190               ElseIf CTDict.Exists(UnQuotedHeader) Then
191                   CT = CTDict.item(UnQuotedHeader)
192               ElseIf CTDict.Exists(0) Then
193                   CT = CTDict.item(0)
194               Else
195                   CT = False
196               End If
                  
197               ParseCTString CT, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ConvertQuoted, TrimFields
                  
198               AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
                      ShowBooleansAsBooleans Or ShowErrorsAsErrors
                      
199               Set Sentinels = New Scripting.Dictionary
                  
200               MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings

201               For i = 1 To NR
202                   If Not IsEmpty(ReturnArray(i + Adj, j + Adj)) Then
203                       Field = CStr(ReturnArray(i + Adj, j + Adj))
204                       QC = CountQuotes(Field, DQ)
205                       ReturnArray(i + Adj, j + Adj) = ConvertField(Field, AnyConversion, _
                              Len(ReturnArray(i + Adj, j + Adj)), TrimFields, DQ, QC, ConvertQuoted, _
                              ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                              ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, _
                              DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, Sentinels, _
                              MaxSentinelLength, ShowMissingsAs)
206                   End If
207               Next i
208           Next j
209       End If

          'Pad if necessary
210       If Not ShowMissingsAsEmpty Then
211           If NumColsInReturn > NumColsFound - SkipToCol + 1 Then
212               For i = 1 To NumRowsInReturn
213                   For j = NumColsFound - SkipToCol + 2 To NumColsInReturn
214                       ReturnArray(i + Adj, j + Adj) = ShowMissingsAs
215                   Next j
216               Next i
217           End If
218           If NumRowsInReturn > NumRowsFound Then
219               For i = NumRowsFound + 1 To NumRowsInReturn
220                   For j = 1 To NumColsInReturn
221                       ReturnArray(i + Adj, j + Adj) = ShowMissingsAs
222                   Next j
223               Next i
224           End If
225       End If

226       sCSVRead = ReturnArray

227       Exit Function

ErrHandler:
228       ErrRet = "#sCSVRead: " & Err.Description & "!"
229       If m_ErrorStyle = es_ReturnString Then
230           sCSVRead = ErrRet
231       Else
232           Throw ErrRet
233       End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function sCSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVRead()
          Const Description As String = "Returns the contents of a comma-separated file on disk as an array."
          Dim ArgDescs() As String

1         On Error GoTo ErrHandler

2         ReDim ArgDescs(1 To 19)
3         ArgDescs(1) = "The full name of the file, including the path, or else a URL of a file, or else a string in CSV " & _
              "format."
4         ArgDescs(2) = "Type conversion: Boolean or string, subset of letters NDBETQ. N = convert Numbers, D = convert " & _
              "Dates, B = Convert Booleans, E = convert Excel errors, T = trim leading & trailing spaces, Q = " & _
              "quoted fields also converted. TRUE = NDB, FALSE = no conversion."
5         ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found " & _
              "outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's " & _
              "contents as would be displayed in a text editor."
6         ArgDescs(4) = "Whether delimiters which appear at the start of a line, the end of a line or immediately after " & _
              "another delimiter should be ignored while parsing; useful for fixed-width files with delimiter " & _
              "padding between fields."
7         ArgDescs(5) = "The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO` " & _
              "for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g. " & _
              "2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time."
8         ArgDescs(6) = "Rows that start with this string will be skipped while parsing."
9         ArgDescs(7) = "Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each column " & _
              "will be assigned ShowMissingsAs for that empty row)."
10        ArgDescs(8) = "The row in the file containing headers. Optional and defaults to 0. Type conversion is not " & _
              "applied to fields in the header row, though leading and trailing spaces are trimmed."
11        ArgDescs(9) = "The first row in the file that's included in the return. Optional and defaults to one more than " & _
              "HeaderRowNum."
12        ArgDescs(10) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the " & _
              "first column."
13        ArgDescs(11) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the " & _
              "end of the file are read."
14        ArgDescs(12) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol " & _
              "are read."
15        ArgDescs(13) = "Indicates how `TRUE` values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings; by default, `TRUE`, `True` and `true` are recognised."
16        ArgDescs(14) = "Indicates how `FALSE` values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings; by default, `FALSE`, `False` and `false` are recognised."
17        ArgDescs(15) = "Indicates how missing values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings. By default, only an empty field (consecutive delimiters) is " & _
              "considered missing."
18        ArgDescs(16) = "Fields which are missing in the file (consecutive delimiters) or match one of the " & _
              "MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null " & _
              "string or `#N/A!` error value can be good alternatives."
19        ArgDescs(17) = "Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be " & _
              "omitted and sCSVRead will detect the file's encoding."
20        ArgDescs(18) = "The character that represents a decimal point. If omitted, then the value from Windows " & _
              "regional settings is used."
21        ArgDescs(19) = "For use from VBA only."
22        Application.MacroOptions "sCSVRead", Description, , , , , , , , , ArgDescs
23        Exit Sub

ErrHandler:
24        Debug.Print "Warning: Registration of function sCSVRead failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferSourceType
' Purpose    : Guess whether FileName is in fact a file, a URL or a string in CSV format
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferSourceType(ByVal FileName As String) As enmSourceType

1         On Error GoTo ErrHandler
2         If InStr(FileName, vbLf) > 0 Then 'vbLf and vbCr are not permitted characters in file names or urls
3             InferSourceType = st_String
4         ElseIf InStr(FileName, vbCr) > 0 Then
5             InferSourceType = st_String
6         ElseIf Mid$(FileName, 2, 2) = ":\" Then
7             InferSourceType = st_File
8         ElseIf Left$(FileName, 2) = "\\" Then
9             InferSourceType = st_File
10        ElseIf Left$(FileName, 8) = "https://" Then
11            InferSourceType = st_URL
12        ElseIf Left$(FileName, 7) = "http://" Then
13            InferSourceType = st_URL
14        Else
              'Doesn't look like either file with path, url or string in CSV format
15            InferSourceType = st_String
16            If Len(FileName) < 1000 Then
17                If FileExists(FileName) Then 'file exists in current working directory
18                    InferSourceType = st_File
19                End If
20            End If
21        End If

22        Exit Function
ErrHandler:
23        Throw "#InferSourceType: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MaxStringLengthInArray
' Purpose    : Different versions of Excel have different limits for the longest string that can be an element of an
'              array passed from a VBA UDF back to Excel. I know the limit is 255 for Excel 2010 & 2013 and earlier, and is
'              32,767 for Excel 365 (as of Sep 2021). But don't yet know the limit for Excel 2016 and 2019.
' Tried to get info from StackOverflow, without much joy:
' https://stackoverflow.com/questions/69303804/excel-versions-and-limits-on-the-length-of-string-elements-in-arrays-returned-by
' -----------------------------------------------------------------------------------------------------------------------
Private Function MaxStringLengthInArray() As Long
          Static Res As Long
1         If Res = 0 Then
2             Select Case Val(Application.Version)
                  Case Is <= 15 'Excel 2013. Tested 14 Dec 2021
3                     Res = 255
4                 Case Else
5                     Res = 32767 'Excel 2016, 2019, 365. Hopefully these versions (which all _
                                   return 16 as Application.Version) have the same limit.
6             End Select
7         End If
8         MaxStringLengthInArray = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Download
' Purpose   : Downloads bits from the Internet and saves them to a file.
'             See https://msdn.microsoft.com/en-us/library/ms775123(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Private Function Download(ByVal URLAddress As String, ByVal FileName As String) As String
          Dim ErrString As String
          Dim Res As Long
          Dim TargetFolder As String
          Dim EN As Long

1         On Error GoTo ErrHandler
          
2         TargetFolder = FileFromPath(FileName, False)
3         CreatePath TargetFolder
4         If FileExists(FileName) Then
5             On Error Resume Next
6             FileDelete FileName
7             EN = Err.Number
8             On Error GoTo ErrHandler
9             If EN <> 0 Then
10                Throw "Cannot download from URL '" + URLAddress + "' because target file '" + FileName + _
                      "' already exists and cannot be deleted. Is the target file open in a program such as Excel?"
11            End If
12        End If
          
13        Res = URLDownloadToFile(0, URLAddress, FileName, 0, 0)
14        If Res <> 0 Then
15            ErrString = ParseDownloadError(CLng(Res))
16            Throw "Windows API function URLDownloadToFile returned error code " & CStr(Res) & _
                  " with description '" & ErrString & "'"
17        End If
18        If Not FileExists(FileName) Then Throw "Windows API function URLDownloadToFile did not report an error, " & _
              "but appears to have not successfuly downloaded a file from " & URLAddress & " to " & FileName
              
19        Download = FileName

20        Exit Function
ErrHandler:
21        Throw "#Download: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDownloadError, sub of Download
'              https://www.vbforums.com/showthread.php?882757-URLDownloadToFile-error-codes
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseDownloadError(ByVal ErrNum As Long) As String
          Dim ErrString As String
1         Select Case ErrNum
              Case &H80004004
2                 ErrString = "Aborted"
3             Case &H800C0001
4                 ErrString = "Destination File Exists"
5             Case &H800C0002
6                 ErrString = "Invalid Url"
7             Case &H800C0003
8                 ErrString = "No Session"
9             Case &H800C0004
10                ErrString = "Cannot Connect"
11            Case &H800C0005
12                ErrString = "Resource Not Found"
13            Case &H800C0006
14                ErrString = "Object Not Found"
15            Case &H800C0007
16                ErrString = "Data Not Available"
17            Case &H800C0008
18                ErrString = "Download Failure"
19            Case &H800C0009
20                ErrString = "Authentication Required"
21            Case &H800C000A
22                ErrString = "No Valid Media"
23            Case &H800C000B
24                ErrString = "Connection Timeout"
25            Case &H800C000C
26                ErrString = "Invalid Request"
27            Case &H800C000D
28                ErrString = "Unknown Protocol"
29            Case &H800C000E
30                ErrString = "Security Problem"
31            Case &H800C000F
32                ErrString = "Cannot Load Data"
33            Case &H800C0010
34                ErrString = "Cannot Instantiate Object"
35            Case &H800C0014
36                ErrString = "Redirect Failed"
37            Case &H800C0015
38                ErrString = "Redirect To Dir"
39            Case &H800C0016
40                ErrString = "Cannot Lock Request"
41            Case Else
42                ErrString = "Unknown"
43        End Select
44        ParseDownloadError = ErrString
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadAllFromStream
' Purpose    : Handles both ADOB.Stream and Scripting.TextStream. Note that ADODB.ReadText(-1) to read all of a stream
'              in a single operation has _very_ poor performance for large files, but reading 10,000 characters at a time
'              in a loop appears to solve that problem.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReadAllFromStream(Stream As Object) As String
            
          Dim Chunk As String
          Dim Contents As String
          Dim i As Long
          Const ChunkSize As Long = 10000

1         On Error GoTo ErrHandler
2         If TypeName(Stream) = "Stream" Then
3             Contents = String(ChunkSize, " ")

4             i = 1
5             Do While Not Stream.EOS
6                 Chunk = Stream.ReadText(ChunkSize)
7                 If i - 1 + Len(Chunk) > Len(Contents) Then
8                     Contents = Contents & String(i - 1 + Len(Chunk), " ")
9                 End If

10                Mid$(Contents, i, Len(Chunk)) = Chunk
11                i = i + Len(Chunk)
12            Loop

13            If (i - 1) < Len(Contents) Then
14                Contents = Left$(Contents, i - 1)
15            End If

16            ReadAllFromStream = Contents
17        ElseIf TypeName(Stream) = "TextStream" Then
18            ReadAllFromStream = Stream.ReadAll
19        Else
20            Throw "Stream has unknown type: " + TypeName(Stream)
21        End If

22        Exit Function
ErrHandler:
23        Throw "#ReadAllFromStream: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseEncoding
' Purpose    : Set by-ref arguments
' Parameters :
'  FileName:
'  Encoding: Optional argument passed in to sCSVRead. If not passed, we delegate to DetectEncoding.
'  TriState: Set by reference. Needed only when we read files using Scripting.TextStream, i.e. when useADODB is False.
'  CharSet : Set by reference. Needed only when we read files using ADODB.Stream, i.e. when useADODB is True.
'  useADODB: Should file be read via ADODB.Stream, which is capable of reading UTF-8 files
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseEncoding(ByVal FileName As String, ByVal Encoding As Variant, ByRef TriState As Long, _
        ByRef CharSet As String, ByRef useADODB As Boolean)

          Const Err_Encoding As String = "Encoding argument can usually be omitted, but otherwise Encoding be " & _
              "either ""ASCII"", ""ANSI"", ""UTF-8"", ""UTF-8-BOM"", ""UTF-16"" or ""UTF-16-BOM"""
          
1         On Error GoTo ErrHandler
2         If IsEmpty(Encoding) Or IsMissing(Encoding) Then
3             DetectEncoding FileName, TriState, CharSet, useADODB
4         ElseIf VarType(Encoding) = vbString Then
5             Select Case UCase$(Replace(Replace(Encoding, "-", vbNullString), " ", vbNullString))
                  Case "ASCII"
6                     CharSet = "ascii" 'not actually relevant, since we won't use ADODB
7                     TriState = TristateFalse
8                     useADODB = False
9                 Case "ANSI"
10                    CharSet = "_autodetect_all" 'not actually relevant, since we won't use ADODB
11                    TriState = TristateFalse
12                    useADODB = False
13                Case "UTF8"
14                    CharSet = "utf-8"
15                    TriState = TristateFalse
16                    useADODB = True 'Use ADODB because Scripting.TextStream can't cope with UTF-8
17                Case "UTF8BOM"
18                    CharSet = "utf-8"
19                    TriState = TristateFalse
20                    useADODB = True 'Use ADODB because Scripting.TextStream can't cope with UTF-8
21                Case "UTF16"
22                    CharSet = "utf-16"
23                    TriState = TristateTrue
24                    useADODB = False
25                Case "UTF16BOM"
26                    CharSet = "utf-16"
27                    TriState = TristateTrue
28                    useADODB = False
29                Case Else
30                    Throw Err_Encoding
31            End Select
32        Else
33            Throw Err_Encoding
34        End If

35        Exit Sub
ErrHandler:
36        Throw "#ParseEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsCTValid
' Purpose    : Is a "Convert Types string" (which can in fact be either a string or a Boolean) valid?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsCTValid(ByVal CT As Variant) As Boolean

          Static rx As VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler
2         If rx Is Nothing Then
3             Set rx = New RegExp
4             With rx
5                 .IgnoreCase = True
6                 .Pattern = "^[NDBETQ]*$"
7                 .Global = False        'Find first match only
8             End With
9         End If

10        If VarType(CT) = vbBoolean Then
11            IsCTValid = True
12        ElseIf VarType(CT) = vbString Then
13            IsCTValid = rx.Test(CT)
14        Else
15            IsCTValid = False
16        End If

17        Exit Function
ErrHandler:
18        Throw "#IsCTValid: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CTsEqual
' Purpose    : Test if two CT strings (strings to define type conversion) are equal, i.e. will have the same effect
' -----------------------------------------------------------------------------------------------------------------------
Private Function CTsEqual(ByVal CT1 As Variant, ByVal CT2 As Variant) As Boolean
1         On Error GoTo ErrHandler
2         If VarType(CT1) = VarType(CT2) Then
3             If CT1 = CT2 Then
4                 CTsEqual = True
5                 Exit Function
6             End If
7         End If
8         CTsEqual = StandardiseCT(CT1) = StandardiseCT(CT2)
9         Exit Function
ErrHandler:
10        Throw "#CTsEqual: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StandardiseCT
' Purpose    : Put a CT string into standard form so that two such can be compared.
' -----------------------------------------------------------------------------------------------------------------------
Private Function StandardiseCT(ByVal CT As Variant) As String
1         On Error GoTo ErrHandler
2         If VarType(CT) = vbBoolean Then
3             If CT Then
4                 StandardiseCT = "BDN"
5             Else
6                 StandardiseCT = vbNullString
7             End If
8             Exit Function
9         ElseIf VarType(CT) = vbString Then
10            StandardiseCT = IIf(InStr(1, CT, "B", vbTextCompare), "B", vbNullString) & _
                  IIf(InStr(1, CT, "D", vbTextCompare), "D", vbNullString) & _
                  IIf(InStr(1, CT, "E", vbTextCompare), "E", vbNullString) & _
                  IIf(InStr(1, CT, "N", vbTextCompare), "N", vbNullString) & _
                  IIf(InStr(1, CT, "Q", vbTextCompare), "Q", vbNullString) & _
                  IIf(InStr(1, CT, "T", vbTextCompare), "T", vbNullString)
11        End If

12        Exit Function
ErrHandler:
13        Throw "#StandardiseCT: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDArrayToTwoDArray
' Purpose    : Convert 1-d array of 2-element 1-d arrays into a 1-based, two-column, 2-d array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OneDArrayToTwoDArray(ByVal x As Variant) As Variant
          Const Err_1DArray As String = "If ConvertTypes is given as a 1-dimensional array, each element must " & _
              "be a 1-dimensional array with two elements"

          Dim i As Long
          Dim k As Long
          Dim TwoDArray() As Variant
1         On Error GoTo ErrHandler
2         ReDim TwoDArray(1 To UBound(x) - LBound(x) + 1, 1 To 2)
3         For i = LBound(x) To UBound(x)
4             k = k + 1
5             If Not IsArray(x(i)) Then Throw Err_1DArray
6             If NumDimensions(x(i)) <> 1 Then Throw Err_1DArray
7             If UBound(x(i)) - LBound(x(i)) <> 1 Then Throw Err_1DArray
8             TwoDArray(k, 1) = x(i)(LBound(x(i)))
9             TwoDArray(k, 2) = x(i)(1 + LBound(x(i)))
10        Next i
11        OneDArrayToTwoDArray = TwoDArray
12        Exit Function
ErrHandler:
13        Throw "#OneDArrayToTwoDArray: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : There is flexibility in how the ConvertTypes argument is provided to sCSVRead:
'              a) As a string or Boolean for the same type conversion rules for every field in the file; or
'              b) An array to define different type conversion rules by column, in which case ConvertTypes can be passed
'                 as a two-column or two-row array (convenient from Excel) or as an array of two-element arrays
'                 (convenient from VBA).
'              If an array, then the left col(or top row or first element) can contain either column numbers or strings
'              that match the elements of the SkipToRow row of the file
'
' Parameters :
'  ConvertTypes          :
'  ShowNumbersAsNumbers  : Set only if ConvertTypes is not an array
'  ShowDatesAsDates      : Set only if ConvertTypes is not an array
'  ShowBooleansAsBooleans: Set only if ConvertTypes is not an array
'  ShowErrorsAsErrors    : Set only if ConvertTypes is not an array
'  ConvertQuoted         : Set only if ConvertTypes is not an array
'  TrimFields            : Set only if ConvertTypes is not an array
'  ColByColFormatting    : Set to True if ConvertTypes is an array
'  HeaderRowNum          : As passed to sCSVRead, used to throw an error if HeaderRowNum has not been specified when
'                          it needs to have been.
'  CTDict                : Set to a dictionary keyed on the elements of the left column (or top row) of ConvertTypes,
'                          each element containing the corresponding right (or bottom) element.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
        ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
        ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean, _
        ByRef ColByColFormatting As Boolean, ByVal HeaderRowNum As Long, ByRef CTDict As Scripting.Dictionary)
          
          Const Err_2D As String = "If ConvertTypes is given as a two dimensional array then the " & _
              " lower bounds in each dimension must be 1"
          Const Err_Ambiguous As String = "ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns"
          Const Err_BadColumnIdentifier As String = "Column identifiers in the left column (or top row) of " & _
              "ConvertTypes must be strings or non-negative whole numbers"
          Const Err_BadCT As String = "Type Conversion given in bottom row (or right column) of ConvertTypes must be " & _
              "Booleans or strings containing letters NDBETQ"
          Const Err_ConvertTypes As String = "ConvertTypes must be a Boolean, a string with allowed letters ""NDBETQ"" or an array"
          Const Err_HeaderRowNum As String = "ConvertTypes specifies columns by their header (instead of by number), " & _
              "but HeaderRowNum has not been specified"
          
          Dim ColIdentifier As Variant
          Dim CT As Variant
          Dim i As Long
          Dim LCN As Long 'Left column number
          Dim NC As Long 'Number of columns
          Dim ND As Long 'Number of dimensions
          Dim NR As Long 'Number of rows
          Dim RCN As Long 'Right Column Number
          Dim Transposed As Boolean
          
1         On Error GoTo ErrHandler
2         If VarType(ConvertTypes) = vbString Or VarType(ConvertTypes) = vbBoolean Or IsEmpty(ConvertTypes) Then
3             ParseCTString CStr(ConvertTypes), ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                  ShowErrorsAsErrors, ConvertQuoted, TrimFields
4             ColByColFormatting = False
5             Exit Sub
6         End If

7         If TypeName(ConvertTypes) = "Range" Then ConvertTypes = ConvertTypes.Value2
8         ND = NumDimensions(ConvertTypes)
9         If ND = 1 Then
10            ConvertTypes = OneDArrayToTwoDArray(ConvertTypes)
11        ElseIf ND = 2 Then
12            If LBound(ConvertTypes, 1) <> 1 Or LBound(ConvertTypes, 2) <> 1 Then
13                Throw Err_2D
14            End If
15        End If

16        NR = NRows(ConvertTypes)
17        NC = NCols(ConvertTypes)
18        If NR = 2 And NC = 2 Then
              'Tricky - have we been given two rows or two columns?
19            If Not IsCTValid(ConvertTypes(2, 2)) Then Throw Err_ConvertTypes
20            If IsCTValid(ConvertTypes(1, 2)) And IsCTValid(ConvertTypes(2, 1)) Then
21                If StandardiseCT(ConvertTypes(1, 2)) <> StandardiseCT(ConvertTypes(2, 1)) Then
22                    Throw Err_Ambiguous
23                End If
24            End If
25            If IsCTValid(ConvertTypes(2, 1)) Then
26                ConvertTypes = Transpose(ConvertTypes)
27                Transposed = True
28            End If
29        ElseIf NR = 2 Then
30            ConvertTypes = Transpose(ConvertTypes)
31            Transposed = True
32            NR = NC
33        ElseIf NC <> 2 Then
34            Throw Err_ConvertTypes
35        End If
36        LCN = LBound(ConvertTypes, 2)
37        RCN = LCN + 1
38        For i = LBound(ConvertTypes, 1) To UBound(ConvertTypes, 1)
39            ColIdentifier = ConvertTypes(i, LCN)
40            CT = ConvertTypes(i, RCN)
41            If IsNumber(ColIdentifier) Then
42                If ColIdentifier <> CLng(ColIdentifier) Then
43                    Throw Err_BadColumnIdentifier & _
                          " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                          ") is " & CStr(ColIdentifier)
44                ElseIf ColIdentifier < 0 Then
45                    Throw Err_BadColumnIdentifier & " but ConvertTypes(" & _
                          IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
46                End If
47            ElseIf VarType(ColIdentifier) <> vbString Then
48                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                      ") is of type " & TypeName(ColIdentifier)
49            End If
50            If Not IsCTValid(CT) Then
51                If VarType(CT) = vbString Then
52                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is string """ & CStr(CT) & """"
53                Else
54                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is of type " & TypeName(CT)
55                End If
56            End If

57            If CTDict.Exists(ColIdentifier) Then
58                If Not CTsEqual(CTDict.item(ColIdentifier), CT) Then
59                    Throw "ConvertTypes is contradictory. Column " & CStr(ColIdentifier) & _
                          " is specified to be converted using two different conversion rules: " & CStr(CT) & _
                          " and " & CStr(CTDict.item(ColIdentifier))
60                End If
61            Else
62                CT = StandardiseCT(CT)
                  'Need this line to ensure that we parse the DateFormat provided when doing Col-by-col type conversion
63                If InStr(CT, "D") > 0 Then ShowDatesAsDates = True
64                If VarType(ColIdentifier) = vbString Then
65                    If HeaderRowNum = 0 Then
66                        Throw Err_HeaderRowNum
67                    End If
68                End If
69                CTDict.Add ColIdentifier, CT
70            End If
71        Next i
72        ColByColFormatting = True
73        Exit Sub
ErrHandler:
74        Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCTString
' Purpose    : Parse the input ConvertTypes to set seven Boolean flags which are passed by reference.
' Parameters :
'  ConvertTypes          : The argument to sCSVRead
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as
'                          Dates?
'  ShowBooleansAsBooleans: Should fields in the file that match one of the TrueStrings or FalseStrings be returned as
'                          Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  ConvertQuoted         : Should the four conversion rules above apply even to quoted fields?
'  TrimFields            : Should leading and trailing spaces be trimmed from fields?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCTString(ByVal ConvertTypes As String, ByRef ShowNumbersAsNumbers As Boolean, _
        ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
        ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean)

          Const Err_ConvertTypes As String = "ConvertTypes must be Boolean or string with allowed letters NDBETQ. " & _
              """N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans " & _
              "as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing " & _
              "spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" " & _
              "(convert unquoted numbers, dates and Booleans), FALSE = no conversion"
          Const Err_Quoted As String = "ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to " & _
              "quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply"
          Dim i As Long

1         On Error GoTo ErrHandler

2         If ConvertTypes = "True" Or ConvertTypes = "False" Then
3             ConvertTypes = StandardiseCT(CBool(ConvertTypes))
4         End If

5         ShowNumbersAsNumbers = False
6         ShowDatesAsDates = False
7         ShowBooleansAsBooleans = False
8         ShowErrorsAsErrors = False
9         ConvertQuoted = False
10        For i = 1 To Len(ConvertTypes)
              'Adding another letter? Also change method IsCTValid!
11            Select Case UCase$(Mid$(ConvertTypes, i, 1))
                  Case "N"
12                    ShowNumbersAsNumbers = True
13                Case "D"
14                    ShowDatesAsDates = True
15                Case "B"
16                    ShowBooleansAsBooleans = True
17                Case "E"
18                    ShowErrorsAsErrors = True
19                Case "Q"
20                    ConvertQuoted = True
21                Case "T"
22                    TrimFields = True
23                Case Else
24                    Throw Err_ConvertTypes & " Found unrecognised character '" _
                          & Mid$(ConvertTypes, i, 1) & "'"
25            End Select
26        Next i
          
27        If ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
28            Throw Err_Quoted
29        End If

30        Exit Sub
ErrHandler:
31        Throw "#ParseCTString: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(ByVal N1 As Long, ByVal N2 As Long, ByVal N3 As Long, _
        ByVal N4 As Long, ByRef Which As Long) As Long

1         If N1 < N2 Then
2             Min4 = N1
3             Which = 1
4         Else
5             Min4 = N2
6             Which = 2
7         End If

8         If N3 < Min4 Then
9             Min4 = N3
10            Which = 3
11        End If

12        If N4 < Min4 Then
13            Min4 = N4
14            Which = 4
15        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DetectEncoding
' Purpose    : Guesses whether a file needs to be opened with the "format" argument to File.OpenAsTextStream set to
'              TriStateTrue or TriStateFalse.
'              The documentation at
'              https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
'              is limited but I believe that:
'            * TriStateTrue needs to passed for files which (as reported by NotePad++) are encoded as either
'              "UTF-16 LE BOM" or "UTF-16 BE BOM"
'            * TristateFalse needs to be passed for files encoded as "ANSI"
'            * UTF-8 files are not correctly handled by OpenAsTextStream, instead we use ADODB.Stream, setting CharSet
'              to "UTF-8".
' -----------------------------------------------------------------------------------------------------------------------
Private Sub DetectEncoding(ByVal FilePath As String, ByRef TriState As Long, _
        ByRef CharSet As String, ByRef useADODB As Boolean)

          Dim intAsc1Chr As Long
          Dim intAsc2Chr As Long
          Dim intAsc3Chr As Long
          Dim t As Scripting.TextStream

1         On Error GoTo ErrHandler
          
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
          
3         If (m_FSO.FileExists(FilePath) = False) Then
4             Throw "File not found!"
5         End If

          ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
6         Set t = m_FSO.OpenTextFile(FilePath, 1, False, 0)
7         If t.atEndOfStream Then
8             t.Close: Set t = Nothing
9             TriState = TristateFalse
10            CharSet = "_autodetect_all"
11            useADODB = False
12            Exit Sub
13        End If
14        intAsc1Chr = Asc(t.Read(1))
15        If t.atEndOfStream Then
16            t.Close: Set t = Nothing
17            TriState = TristateFalse
18            CharSet = "_autodetect_all"
19            useADODB = False
20            Exit Sub
21        End If
          
22        intAsc2Chr = Asc(t.Read(1))
          
23        If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
              'File is probably encoded UTF-16 LE BOM (little endian, with Byte Option Marker)
24            TriState = TristateTrue
25            CharSet = "utf-16"
26            useADODB = False
27        ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
              'File is probably encoded UTF-16 BE BOM (big endian, with Byte Option Marker)
28            TriState = TristateTrue
29            CharSet = "utf-16"
30            useADODB = False
31        Else
32            If t.atEndOfStream Then
33                TriState = TristateFalse
34                Exit Sub
35            End If
36            intAsc3Chr = Asc(t.Read(1))
37            If (intAsc1Chr = 239) And (intAsc2Chr = 187) And (intAsc3Chr = 191) Then
                  'File is probably encoded UTF-8 with BOM
38                CharSet = "utf-8"
39                TriState = TristateFalse
40                useADODB = True
41            Else
                  'We don't know, assume ANSI but that may be incorrect.
42                CharSet = vbNullString
43                TriState = TristateFalse
44                useADODB = False
45            End If
46        End If

47        t.Close: Set t = Nothing
48        Exit Sub
ErrHandler:
49        Throw "#DetectEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|). Only look in the first 10,000 characters, Would prefer to look at the first
'              10 lines, but that presents a problem for files with Mac line endings as T.ReadLine doesn't work for them.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(ByVal st As enmSourceType, ByVal FileNameOrContents As String, _
        ByVal TriState As Long, ByVal DecimalSeparator As String) As String
          
          Const CHUNK_SIZE As Long = 1000
          Const Err_SourceType As String = "Cannot infer delimiter directly from URL"
          Const MAX_CHUNKS As Long = 10
          Const QuoteChar As String = """"
          Dim Contents As String
          Dim CopyOfErr As String
          Dim EvenQuotes As Boolean
          Dim F As Scripting.file
          Dim i As Long
          Dim j As Long
          Dim MaxChars As Long
          Dim t As TextStream

1         On Error GoTo ErrHandler

2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

3         EvenQuotes = True
4         If st = st_File Then

5             Set F = m_FSO.GetFile(FileNameOrContents)
6             Set t = F.OpenAsTextStream(ForReading, TriState)

7             If t.atEndOfStream Then
8                 t.Close: Set t = Nothing: Set F = Nothing
9                 Throw "File is empty"
10            End If

11            Do While Not t.atEndOfStream And j <= MAX_CHUNKS
12                j = j + 1
13                Contents = t.Read(CHUNK_SIZE)
14                For i = 1 To Len(Contents)
15                    Select Case Mid$(Contents, i, 1)
                          Case QuoteChar
16                            EvenQuotes = Not EvenQuotes
17                        Case ",", vbTab, "|", ";", ":"
18                            If EvenQuotes Then
19                                If Mid$(Contents, i, 1) <> DecimalSeparator Then
20                                    InferDelimiter = Mid$(Contents, i, 1)
21                                    t.Close: Set t = Nothing: Set F = Nothing
22                                    Exit Function
23                                End If
24                            End If
25                    End Select
26                Next i
27            Loop
28            t.Close
29        ElseIf st = st_String Then
30            Contents = FileNameOrContents
31            MaxChars = MAX_CHUNKS * CHUNK_SIZE
32            If MaxChars > Len(Contents) Then MaxChars = Len(Contents)

33            For i = 1 To MaxChars
34                Select Case Mid$(Contents, i, 1)
                      Case QuoteChar
35                        EvenQuotes = Not EvenQuotes
36                    Case ",", vbTab, "|", ";", ":"
37                        If EvenQuotes Then
38                            If Mid$(Contents, i, 1) <> DecimalSeparator Then
39                                InferDelimiter = Mid$(Contents, i, 1)
40                                Exit Function
41                            End If
42                        End If
43                End Select
44            Next i
45        Else
46            Throw Err_SourceType
47        End If

          'No commonly-used delimiter found in the file outside quoted regions _
           and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma _
           unless that's the decimal separator.
          
48        If DecimalSeparator = "," Then
49            InferDelimiter = ";"
50        Else
51            InferDelimiter = ","
52        End If

53        Exit Function
ErrHandler:
54        CopyOfErr = "#InferDelimiter: " & Err.Description & "!"
55        If Not t Is Nothing Then
56            t.Close
57            Set t = Nothing: Set F = Nothing
58        End If
59        Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Purpose    : Populate DateOrder and DateSeparator by parsing DateFormat.
' Parameters :
'  DateFormat   : String such as D/M/Y or Y-M-D
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String, _
        ByRef ISO8601 As Boolean, ByRef AcceptWithoutTimeZone As Boolean, ByRef AcceptWithTimeZone As Boolean)

          Dim Err_DateFormat As String

1         On Error GoTo ErrHandler
          
2         If UCase$(DateFormat) = "ISO" Then
3             ISO8601 = True
4             AcceptWithoutTimeZone = True
5             AcceptWithTimeZone = False
6             Exit Sub
7         ElseIf UCase$(DateFormat) = "ISOZ" Then
8             ISO8601 = True
9             AcceptWithoutTimeZone = False
10            AcceptWithTimeZone = True
11            Exit Sub
12        End If
          
13        Err_DateFormat = "DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', " & _
              "'M/D/Y', 'D/M/Y' or 'Y/M/D'" & ". Omit to use the default date format of 'Y-M-D'"
           
          'Replace repeated D's with a single D, etc since CastToDate only needs _
           to know the order in which the three parts of the date appear.
14        If Len(DateFormat) > 5 Then
15            DateFormat = UCase$(DateFormat)
16            ReplaceRepeats DateFormat, "D"
17            ReplaceRepeats DateFormat, "M"
18            ReplaceRepeats DateFormat, "Y"
19        End If
             
20        If Len(DateFormat) = 0 Then 'use "Y-M-D"
21            DateOrder = 2
22            DateSeparator = "-"
23        ElseIf Len(DateFormat) <> 5 Then
24            Throw Err_DateFormat
25        ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
26            Throw Err_DateFormat
27        Else
28            DateSeparator = Mid$(DateFormat, 2, 1)
29            If DateSeparator <> "/" And DateSeparator <> "-" Then Throw Err_DateFormat
30            Select Case UCase$(Left$(DateFormat, 1) & Mid$(DateFormat, 3, 1) & Right$(DateFormat, 1))
                  Case "MDY"
31                    DateOrder = 0
32                Case "DMY"
33                    DateOrder = 1
34                Case "YMD"
35                    DateOrder = 2
36                Case Else
37                    Throw Err_DateFormat
38            End Select
39        End If

40        Exit Sub
ErrHandler:
41        Throw "#ParseDateFormat: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, ByVal TheChar As String)
          Dim ChCh As String
1         ChCh = TheChar & TheChar
2         Do While InStr(TheString, ChCh) > 0
3             TheString = Replace(TheString, ChCh, TheChar)
4         Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting
'              Buffer into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  useADODB        : Pass as True when ContentsOrStream is ADODB.Stream, False when it's Scripting.TextStream,
'                    ignored when it's a string.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  Comment         : Lines in the file that start with these characters will be ignored, handled by method SkipLines.
'  IgnoreRepeated  : If true then parsing ignores delimiters at the start of lines, consecutive delimiters and delimiters
'                    at the end of lines.
'  SkipToRow       : The first line of the file to appear in the return from sCSVRead. However, we need to parse earlier
'                    lines to identify where SkipToRow starts in the file - see variable HaveReachedSkipToRow.
'  HeaderRowNum    : The row number of the headers in the file, must be less than or equal to SkipToRow.
'  NumRows         : The number of rows to parse. 0 for all rows from SkipToRow to the end of the file.
'  NumRowsFound    : Set to the number of rows in the file that are on or after SkipToRow.
'  NumColsFound    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields       : Set to the number of fields in the file that are on or after SkipToRow.  May be less than
'                    NumRowsFound times NumColsFound if not all lines have the same number of fields.
'  Ragged          : Set to True if not all rows of the file have the same number of fields.
'  Starts          : Set to an array of size at least NumFields. Element k gives the point in Buffer at which the
'                    kth field starts.
'  Lengths         : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  RowIndexes      : Set to an array of size at least NumFields. Element k gives the row at which the kth field should
'                    appear in the return from sCSVRead.
'  ColIndexes      : Set to an array of size at least NumFields. Element k gives the column at which the kth field would
'                    appear in the return from sCSVRead under the assumption that argument SkipToCol is 1.
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in
'                    the kth field.
'  HeaderRow       : Set equal to the contents of the header row in the file, no type conversion, but quoted fields are
'                    unquoted and leading and trailing spaces are removed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, ByVal useADODB As Boolean, ByVal QuoteChar As String, _
        ByVal Delimiter As String, ByVal Comment As String, ByVal IgnoreEmptyLines As Boolean, _
        ByVal IgnoreRepeated As Boolean, ByVal SkipToRow As Long, ByVal HeaderRowNum As Long, ByVal NumRows As Long, _
        ByRef NumRowsFound As Long, ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Ragged As Boolean, _
        ByRef Starts() As Long, ByRef Lengths() As Long, ByRef RowIndexes() As Long, ByRef ColIndexes() As Long, _
        ByRef QuoteCounts() As Long, ByRef HeaderRow As Variant) As String

          Const Err_Delimiter As String = "Delimiter must not be the null string"
          Dim Buffer As String
          Dim BufferUpdatedTo As Long
          Dim ColNum As Long
          Dim DoSkipping As Boolean
          Dim EvenQuotes As Boolean
          Dim HaveReachedSkipToRow As Boolean
          Dim i As Long 'Index to read from Buffer
          Dim j As Long 'Index to write to Starts, Lengths, RowIndexes and ColIndexes
          Dim LComment As Long
          Dim LDlm As Long
          Dim NumRowsInFile As Long
          Dim OrigLen As Long
          Dim PosCR As Long
          Dim PosDL As Long
          Dim PosLF As Long
          Dim PosQC As Long
          Dim QuoteArray() As String
          Dim quoteCount As Long
          Dim RowNum As Long
          Dim SearchFor() As String
          Dim Stream As Object
          Dim Streaming As Boolean
          Dim tmp As Long
          Dim Which As Long

1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
3         HeaderRow = Empty
          
4         If VarType(ContentsOrStream) = vbString Then
5             Buffer = ContentsOrStream
6             Streaming = False
7         Else
8             Set Stream = ContentsOrStream
9             If NumRows = 0 Then
10                Buffer = ReadAllFromStream(Stream)
11                Streaming = False
12            Else
13                GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
14                Streaming = True
15            End If
16        End If
             
17        LComment = Len(Comment)
18        If LComment > 0 Or IgnoreEmptyLines Then
19            DoSkipping = True
20        End If
             
21        If Streaming Then
22            ReDim SearchFor(1 To 4)
23            SearchFor(1) = Delimiter
24            SearchFor(2) = vbLf
25            SearchFor(3) = vbCr
26            SearchFor(4) = QuoteChar
27            ReDim QuoteArray(1 To 1)
28            QuoteArray(1) = QuoteChar
29        End If

30        ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
31        ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
          
32        LDlm = Len(Delimiter)
33        If LDlm = 0 Then Throw Err_Delimiter 'Avoid infinite loop!
34        OrigLen = Len(Buffer)
35        If Not Streaming Then
              'Ensure Buffer terminates with vbCrLf
36            If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
37                Buffer = Buffer & vbCrLf
38            ElseIf Right$(Buffer, 1) = vbCr Then
39                Buffer = Buffer & vbLf
40            End If
41            BufferUpdatedTo = Len(Buffer)
42        End If
          
43        i = 0: j = 1
          
44        If DoSkipping Then
45            SkipLines Streaming, useADODB, Comment, LComment, IgnoreEmptyLines, _
                  Stream, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
46        End If
          
47        If IgnoreRepeated Then
              'IgnoreRepeated: Handle repeated delimiters at the start of the first line
48            Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
49                i = i + LDlm
50            Loop
51        End If
          
52        ColNum = 1: RowNum = 1
53        EvenQuotes = True
54        Starts(1) = i + 1
55        If SkipToRow = 1 Then HaveReachedSkipToRow = True

56        Do
57            If EvenQuotes Then
58                If Not Streaming Then
59                    If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
60                    If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
61                    If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
62                    If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
63                    i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
64                Else
65                    i = SearchInBuffer(SearchFor, i + 1, Stream, useADODB, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
66                End If

67                If i >= BufferUpdatedTo + 1 Then
68                    Exit Do
69                End If

70                If j + 1 > UBound(Starts) Then
71                    ReDim Preserve Starts(1 To UBound(Starts) * 2)
72                    ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
73                    ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
74                    ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
75                    ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
76                End If

77                Select Case Which
                      Case 1
                          'Found Delimiter
78                        Lengths(j) = i - Starts(j)
79                        If IgnoreRepeated Then
80                            Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
81                                i = i + LDlm
82                            Loop
83                        End If
                          
84                        Starts(j + 1) = i + LDlm
85                        ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
86                        ColNum = ColNum + 1
87                        QuoteCounts(j) = quoteCount: quoteCount = 0
88                        j = j + 1
89                        NumFields = NumFields + 1
90                        i = i + LDlm - 1
91                    Case 2, 3
                          'Found line ending
92                        Lengths(j) = i - Starts(j)
93                        If Which = 3 Then 'Found a vbCr
94                            If Mid$(Buffer, i + 1, 1) = vbLf Then
                                  'Ending is Windows rather than Mac or Unix.
95                                i = i + 1
96                            End If
97                        End If
                          
98                        If DoSkipping Then
99                            SkipLines Streaming, useADODB, Comment, LComment, IgnoreEmptyLines, Stream, _
                                  Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
100                       End If
                          
101                       If IgnoreRepeated Then
                              'IgnoreRepeated: Handle repeated delimiters at the end of the line, _
                               all but one will have already been handled.
102                           If Lengths(j) = 0 Then
103                               If ColNum > 1 Then
104                                   j = j - 1
105                                   ColNum = ColNum - 1
106                                   NumFields = NumFields - 1
107                               End If
108                           End If
                              'IgnoreRepeated: handle delimiters at the start of the next line to be parsed
109                           Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
110                               i = i + LDlm
111                           Loop
112                       End If
113                       Starts(j + 1) = i + 1

114                       If ColNum > NumColsFound Then
115                           If NumColsFound > 0 Then
116                               Ragged = True
117                           End If
118                           NumColsFound = ColNum
119                       ElseIf ColNum < NumColsFound Then
120                           Ragged = True
121                       End If
                          
122                       ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
123                       QuoteCounts(j) = quoteCount: quoteCount = 0
                          
124                       If RowNum = 1 Then
125                           If SkipToRow = 1 Then
126                               If HeaderRowNum = 1 Then
127                                   HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, _
                                          ColIndexes, QuoteCounts, j)
128                               End If
129                           End If
130                       End If
131                       If Not HaveReachedSkipToRow Then
132                           If RowNum = HeaderRowNum Then
133                               HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, _
                                      ColIndexes, QuoteCounts, j)
134                           End If
135                       End If
                          
136                       ColNum = 1: RowNum = RowNum + 1
                          
137                       j = j + 1
138                       NumFields = NumFields + 1
                          
139                       If HaveReachedSkipToRow Then
140                           If RowNum = NumRows + 1 Then
141                               Exit Do
142                           End If
143                       Else
144                           If RowNum = SkipToRow Then
145                               HaveReachedSkipToRow = True
146                               tmp = Starts(j)
147                               ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
148                               ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
149                               RowNum = 1: j = 1: NumFields = 0
150                               Starts(1) = tmp
151                           End If
152                       End If
153                   Case 4
                          'Found QuoteChar
154                       EvenQuotes = False
155                       quoteCount = quoteCount + 1
156               End Select
157           Else
158               If Not Streaming Then
159                   PosQC = InStr(i + 1, Buffer, QuoteChar)
160               Else
161                   If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray, i + 1, Stream, useADODB, _
                          Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
162               End If
                  
163               If PosQC = 0 Then
                      'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                       If there are an odd number then all text after the last double quote in the file will be (part of) _
                       the last field in the last line.
164                   Lengths(j) = OrigLen - Starts(j) + 1
165                   ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                      
166                   RowNum = RowNum + 1
167                   If ColNum > NumColsFound Then NumColsFound = ColNum
168                   NumFields = NumFields + 1
169                   Exit Do
170               Else
171                   i = PosQC
172                   EvenQuotes = True
173                   quoteCount = quoteCount + 1
174               End If
175           End If
176       Loop

177       NumRowsFound = RowNum - 1
          
178       If HaveReachedSkipToRow Then
179           NumRowsInFile = SkipToRow - 1 + RowNum - 1
180       Else
181           NumRowsInFile = RowNum - 1
182       End If
          
183       If SkipToRow > NumRowsInFile Then
184           If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or _
                                   a negative number of rows. So throw an error.
                  Dim RowDescription As String
185               If IgnoreEmptyLines And Len(Comment) > 0 Then
186                   RowDescription = "not commented, not empty "
187               ElseIf IgnoreEmptyLines Then
188                   RowDescription = "not empty "
189               ElseIf Len(Comment) > 0 Then
190                   RowDescription = "not commented "
191               End If
                                   
192               Throw "SkipToRow (" & CStr(SkipToRow) & ") exceeds the number of " & RowDescription & _
                      "rows in the file (" & CStr(NumRowsInFile) & ")"
193           Else
                  'Attempting to read a set number of rows, function sCSVRead will return an array of Empty values.
194               NumFields = 0
195               NumRowsFound = 0
196           End If
197       End If

198       ParseCSVContents = Buffer

199       Exit Function
ErrHandler:
200       Throw "#ParseCSVContents: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLastParsedRow
' Purpose    : For use during parsing (fn ParseCSVContents) to grab the header row (which may or may not be part of the
'              function return). The argument j should point into the Starts, Lengths etc arrays, pointing to the last
'              field in the header row
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLastParsedRow(Buffer As String, Starts() As Long, Lengths() As Long, _
        ColIndexes() As Long, QuoteCounts() As Long, j As Long) As Variant
          Dim NC As Long

          Dim Field As String
          Dim i As Long
          Dim Res() As String

1         On Error GoTo ErrHandler
2         NC = ColIndexes(j)

3         ReDim Res(1 To 1, 1 To NC)
4         For i = j To j - NC + 1 Step -1
5             Field = Mid$(Buffer, Starts(i), Lengths(i))
6             Res(1, NC + i - j) = Unquote(Trim$(Field), """", QuoteCounts(i))
7         Next i
8         GetLastParsedRow = Res

9         Exit Function
ErrHandler:
10        Throw "#GetLastParsedRow: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SkipLines
' Purpose    : Sub-routine of ParseCSVContents. Skip a commented or empty row by incrementing i to the position of
'              the line feed just before the next not-commented line.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SkipLines(ByVal Streaming As Boolean, ByVal useADODB As Boolean, ByVal Comment As String, _
        ByVal LComment As Long, ByVal IgnoreEmptyLines As Boolean, Stream As Object, ByVal Delimiter As String, _
        ByRef Buffer As String, ByRef i As Long, ByVal QuoteChar As String, ByVal PosLF As Long, ByVal PosCR As Long, _
        ByRef BufferUpdatedTo As Long)
          
          Dim atEndOfStream As Boolean
          Dim LookAheadBy As Long
          Dim SkipThisLine As Boolean
1         On Error GoTo ErrHandler
2         Do
3             If Streaming Then
4                 LookAheadBy = MaxLngs(LComment, 2)
5                 If i + LookAheadBy > BufferUpdatedTo Then
6                     If useADODB Then
7                         atEndOfStream = Stream.EOS
8                     Else
9                         atEndOfStream = Stream.atEndOfStream
10                    End If
11                    If Not atEndOfStream Then
12                        GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
13                    End If
14                End If
15            End If

16            SkipThisLine = False
17            If LComment > 0 Then
18                If Mid$(Buffer, i + 1, LComment) = Comment Then
19                    SkipThisLine = True
20                End If
21            End If
22            If IgnoreEmptyLines Then
23                Select Case Mid$(Buffer, i + 1, 1)
                      Case vbLf, vbCr
24                        SkipThisLine = True
25                End Select
26            End If

27            If SkipThisLine Then
28                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
29                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
30                If PosLF < PosCR Then
31                    i = PosLF
32                ElseIf PosLF = PosCR + 1 Then
33                    i = PosLF
34                Else
35                    i = PosCR
36                End If
37            Else
38                Exit Do
39            End If
40        Loop

41        Exit Sub
ErrHandler:
42        Throw "#SkipLines: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Sub-routine of ParseCSVContents. Returns the location in the buffer of the first-encountered string
'              amongst the elements of SearchFor, starting the search at point SearchFrom and finishing the search at
'              point BufferUpdatedTo. If none found in that region returns BufferUpdatedTo + 1. Otherwise returns the
'              location of the first found and sets the by-reference argument Which to indicate which element of
'              SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, ByVal StartingAt As Long, Stream As Object, _
        ByVal useADODB As Boolean, ByVal Delimiter As String, ByVal QuoteChar As String, ByRef Which As Long, _
        ByRef Buffer As String, ByRef BufferUpdatedTo As Long) As Long

          Dim atEndOfStream As Boolean
          Dim InstrRes As Long
          Dim PrevBufferUpdatedTo As Long

1         On Error GoTo ErrHandler

          'in this call only search as far as BufferUpdatedTo
2         InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
3         If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
4             SearchInBuffer = InstrRes
5             Exit Function
6         Else
7             If useADODB Then
8                 atEndOfStream = Stream.EOS
9             Else
10                atEndOfStream = Stream.atEndOfStream
11            End If
12            If atEndOfStream Then
13                SearchInBuffer = BufferUpdatedTo + 1
14                Exit Function
15            End If
16        End If

17        Do
18            PrevBufferUpdatedTo = BufferUpdatedTo
19            GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
20            InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
21            If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
22                SearchInBuffer = InstrRes
23                Exit Function
24            ElseIf Stream.EOS Then
25                SearchInBuffer = BufferUpdatedTo + 1
26                Exit Function
27            End If
28        Loop
29        Exit Function

30        Exit Function
ErrHandler:
31        Throw "#SearchInBuffer: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Sub-routine of ParseCSVContents. Returns the first point in SearchWithin at which one of the elements of
'              SearchFor is found, search is restricted to region [StartingAt, EndingAt] and Which is updated with the
'              index identifying which was the first of the strings to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InStrMulti(SearchFor() As String, SearchWithin As String, ByVal StartingAt As Long, _
        ByVal EndingAt As Long, ByRef Which As Long) As Long

          Const Inf As Long = 2147483647
          Dim i As Long
          Dim InstrRes() As Long
          Dim LB As Long
          Dim Result As Long
          Dim UB As Long

1         On Error GoTo ErrHandler
2         LB = LBound(SearchFor): UB = UBound(SearchFor)

3         Result = Inf

4         ReDim InstrRes(LB To UB)
5         For i = LB To UB
6             InstrRes(i) = InStr(StartingAt, SearchWithin, SearchFor(i))
7             If InstrRes(i) > 0 Then
8                 If InstrRes(i) <= EndingAt Then
9                     If InstrRes(i) < Result Then
10                        Result = InstrRes(i)
11                        Which = i
12                    End If
13                End If
14            End If
15        Next
16        InStrMulti = IIf(Result = Inf, 0, Result)

17        Exit Function
ErrHandler:
18        Throw "#InStrMulti: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromStream, Sub-routine of ParseCSVContents
' Purpose    : Write CHUNKSIZE characters from the TextStream T into the buffer, modifying the passed-by-reference
'              arguments  Buffer, BufferUpdatedTo and Streaming.
'              Complexities:
'           a) We have to be careful not to update the buffer to a point part-way through a two-character end-of-line
'              or a multi-character delimiter, otherwise calling method SearchInBuffer might give the wrong result.
'           b) We update a few characters of the buffer beyond the BufferUpdatedTo point with the delimiter, the
'              QuoteChar and vbCrLf. This ensures that the calls to Instr that search the buffer for these strings do
'              not needlessly scan the unupdated part of the buffer.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GetMoreFromStream(t As Variant, ByVal Delimiter As String, ByVal QuoteChar As String, _
    ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

          Const ChunkSize As Long = 5000  ' The number of characters to read from the stream on each call. _
                                            Set to a small number for testing logic and a bigger number for _
                                            performance, but not too high since a common use case is reading _
                                            just the first line of a file. Suggest 5000? Note that when reading _
                                            an entire file (NumRows argument to sCSVRead is zero) function _
                                            GetMoreFromStream is not called.
          Dim atEndOfStream As Boolean
          Dim ExpandBufferBy As Long
          Dim FirstPass As Boolean
          Dim i As Long
          Dim IsScripting As Boolean
          Dim NCharsToWriteToBuffer As Long
          Dim NewChars As String
          Dim OKToExit As Boolean

1         On Error GoTo ErrHandler
          
2         Select Case TypeName(t)
              Case "TextStream"
3                 IsScripting = True
4             Case "Stream"
5                 IsScripting = False
6             Case Else
7                 Throw "T must be of type Scripting.TextStream or ADODB.Stream"
8         End Select
          
9         FirstPass = True
10        Do
11            If IsScripting Then
12                NewChars = t.Read(IIf(FirstPass, ChunkSize, 1))
13                atEndOfStream = t.atEndOfStream
14            Else
15                NewChars = t.ReadText(IIf(FirstPass, ChunkSize, 1))
16                atEndOfStream = t.EOS
17            End If
18            FirstPass = False
19            If atEndOfStream Then
                  'Ensure NewChars terminates with vbCrLf
20                If Right$(NewChars, 1) <> vbCr And Right$(NewChars, 1) <> vbLf Then
21                    NewChars = NewChars & vbCrLf
22                ElseIf Right$(NewChars, 1) = vbCr Then
23                    NewChars = NewChars & vbLf
24                End If
25            End If

26            NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

27            If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
28                ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
29                Buffer = Buffer & String(ExpandBufferBy, "?")
30            End If
              
31            Mid$(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
32            BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

33            OKToExit = True
              'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
34            If Right$(NewChars, 1) = vbCr Then
35                OKToExit = False
36            End If
              'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
37            If Len(Delimiter) > 1 Then
38                For i = 1 To Len(Delimiter) - 1
39                    If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left$(Delimiter, i) Then
40                        OKToExit = False
41                        Exit For
42                    End If
43                Next i
44                If Mid$(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
45                    OKToExit = True
46                End If
47            End If
48            If OKToExit Then Exit Do
49        Loop

          'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
           we can be sure that there is space in the buffer to write the extra characters thanks to
50        Mid$(Buffer, BufferUpdatedTo + 1, 2 + Len(QuoteChar) + Len(Delimiter)) = vbCrLf & QuoteChar & Delimiter

51        Exit Sub
ErrHandler:
52        Throw "#GetMoreFromStream: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CountQuotes
' Purpose    : Count the quotes in a string, only used when applying column-by-column type conversion, because in that
'              case it's not possible to use the count of quotes made at parsing time which is organised row-by-row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CountQuotes(ByVal Str As String, ByVal QuoteChar As String) As Long
          Dim N As Long
          Dim pos As Long

1         Do
2             pos = InStr(pos + 1, Str, QuoteChar)
3             If pos = 0 Then
4                 CountQuotes = N
5                 Exit Function
6             End If
7             N = N + 1
8         Loop
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConvertField
' Purpose    : Convert a field in the file into an element of the returned array.
' Parameters :
'General
'  Field                : Field, i.e. characters from the file between successive delimiters.
'  AnyConversion        : Is any type conversion to take place? i.e. processing other than trimming whitespace and
'                         unquoting.
'  FieldLength          : The length of Field.
'Whitespace and Quotes
'  TrimFields           : Should leading and trailing spaces be trimmed from fields? For quoted fields, this will not
'                         remove spaces between the quotes.
'  QuoteChar            : The quote character, typically ". No support for different opening and closing quote characters
'                         or different escape character.
'  QuoteCount           : How many quote characters does Field contain?
'  ConvertQuoted        : Should quoted fields (after quote removal) be converted according to arguments
'                         ShowNumbersAsNumbers, ShowDatesAsDates, and the contents of Sentinels.
'Numbers
'  ShowNumbersAsNumbers : If Field is a string representation of a number should the function return that number?
'  SepStandard          : Is the decimal separator the same as the system defaults? If True then the next two arguments
'                         are ignored.
'  DecimalSeparator     : The decimal separator used in Field.
'  SysDecimalSeparator  : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates     : If Field is a string representation of a date should the function return that date?
'  ISO8601              : If Field is a date, does it respect (a subset of) ISO8601?
'  AcceptWithoutTimeZone: In the case of ISO8601 dates, should conversion be applied to dates-with-time that have no time
'                         zone information?
'  AcceptWithTimeZone   : In the case of ISO8601 dates, should conversion be applied to dates-with-time that have time
'                         zone information?
'  DateOrder            : If Field is a string representation what order of parts must it respect (not relevant if
'                         ISO8601 is True) 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator        : The date separator, must be either "-" or "/".
'  SysDateOrder         : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator     : The Windows system date separator.
'Booleans, Errors, Missings
'  AnySentinels         : Does the sentinel dictionary have any elements?
'  Sentinels            : A dictionary of Sentinels. If Sentinels.Exists(Field) Then ConvertField = Sentinels(Field)
'  MaxSentinelLength    : The maximum length of the keys of Sentinels.
'  ShowMissingsAs       : The value to which missing fields (consecutive delimiters) are converted. If sCSVRead has a
'                         MissingStrings argument then values matching those strings are also converted to
'                         ShowMissingsAs, thanks to method MakeSentinels.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConvertField(ByVal Field As String, ByVal AnyConversion As Boolean, ByVal FieldLength As Long, _
        ByVal TrimFields As Boolean, ByVal QuoteChar As String, ByVal quoteCount As Long, ByVal ConvertQuoted As Boolean, _
        ByVal ShowNumbersAsNumbers As Boolean, ByVal SepStandard As Boolean, ByVal DecimalSeparator As String, _
        ByVal SysDecimalSeparator As String, ByVal ShowDatesAsDates As Boolean, ByVal ISO8601 As Boolean, _
        ByVal AcceptWithoutTimeZone As Boolean, ByVal AcceptWithTimeZone As Boolean, ByVal DateOrder As Long, _
        ByVal DateSeparator As String, ByVal SysDateOrder As Long, ByVal SysDateSeparator As String, _
        ByVal AnySentinels As Boolean, ByVal Sentinels As Dictionary, ByVal MaxSentinelLength As Long, _
        ByVal ShowMissingsAs As Variant) As Variant

          Dim Converted As Boolean
          Dim dblResult As Double
          Dim dtResult As Date

1         If TrimFields Then
2             If Left$(Field, 1) = " " Then
3                 Field = Trim$(Field)
4                 FieldLength = Len(Field)
5             ElseIf Right$(Field, 1) = " " Then
6                 Field = Trim$(Field)
7                 FieldLength = Len(Field)
8             End If
9         End If

10        If FieldLength = 0 Then
11            ConvertField = ShowMissingsAs
12            Exit Function
13        End If

14        If Not AnyConversion Then
15            If quoteCount = 0 Then
16                ConvertField = Field
17                Exit Function
18            End If
19        End If

20        If quoteCount > 0 Then
21            If Left$(Field, 1) = QuoteChar Then
22                If Right$(Field, 1) = QuoteChar Then
23                    Field = Mid$(Field, 2, FieldLength - 2)
24                    If quoteCount > 2 Then
25                        Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
26                    End If
27                    If ConvertQuoted Then
28                        FieldLength = Len(Field)
29                    Else
30                        ConvertField = Field
31                        Exit Function
32                    End If
33                End If
34            End If
35        End If

36        If AnySentinels Then
37            If FieldLength <= MaxSentinelLength Then
38                If Sentinels.Exists(Field) Then
39                    ConvertField = Sentinels.item(Field)
40                    Exit Function
41                End If
42            End If
43        End If

44        If Not ConvertQuoted Then
45            If quoteCount > 0 Then
46                ConvertField = Field
47                Exit Function
48            End If
49        End If

50        If ShowNumbersAsNumbers Then
51            CastToDouble Field, dblResult, SepStandard, DecimalSeparator, SysDecimalSeparator, Converted
52            If Converted Then
53                ConvertField = dblResult
54                Exit Function
55            End If
56        End If

57        If ShowDatesAsDates Then
58            If ISO8601 Then
59                CastISO8601 Field, dtResult, Converted, AcceptWithoutTimeZone, AcceptWithTimeZone
60            Else
61                CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
62            End If
63            If Not Converted Then
64                If InStr(Field, ":") > 0 Then
65                    CastToTime Field, dtResult, Converted
66                    If Not Converted Then
67                        CastToTimeB Field, dtResult, Converted
68                    End If
69                End If
70            End If
71            If Converted Then
72                ConvertField = dtResult
73                Exit Function
74            End If
75        End If

76        ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unquote
' Purpose    : Unquote a field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Unquote(ByVal Field As String, ByVal QuoteChar As String, ByVal quoteCount As Long) As String

1         On Error GoTo ErrHandler
2         If quoteCount > 0 Then
3             If Left$(Field, 1) = QuoteChar Then
4                 If Right$(QuoteChar, 1) = QuoteChar Then
5                     Field = Mid$(Field, 2, Len(Field) - 2)
6                     If quoteCount > 2 Then
7                         Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
8                     End If
9                 End If
10            End If
11        End If
12        Unquote = Field

13        Exit Function
ErrHandler:
14        Throw "#Unquote: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble, sub-routine of ConvertField
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(ByVal strIn As String, ByRef dblOut As Double, ByVal SepStandard As Boolean, _
        ByVal DecimalSeparator As String, ByVal SysDecimalSeparator As String, ByRef Converted As Boolean)
          
1         On Error GoTo ErrHandler
2         If SepStandard Then
3             dblOut = CDbl(strIn)
4         Else
5             dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
6         End If
7         Converted = True
ErrHandler:
          'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_CastToDate
' Purpose    : Helps tune CastToDate for speed.
'Example output:
'Running SpeedTest_CastToDate 2021-Sep-07 13:36:55
'SysDateOrder = 0
'SysDateSeparator = /
'N = 1,000,000
'Calls per second = 4,733,596  strIn = "foo"                      DateOrder = 2  Result as expected? True
'Calls per second = 4,005,577  strIn = "foo-bar"                  DateOrder = 2  Result as expected? True
'Calls per second = 771,609    strIn = "09-07-2021"               DateOrder = 0  Result as expected? True
'Calls per second = 500,064    strIn = "07-09-2021"               DateOrder = 1  Result as expected? True
'Calls per second = 729,058    strIn = "2021-09-07"               DateOrder = 2  Result as expected? True
'Calls per second = 379,716    strIn = "08-24-2021 15:18:01"      DateOrder = 0  Result as expected? True
'Calls per second = 202,723    strIn = "08-24-2021 15:18:01.123"  DateOrder = 0  Result as expected? True
'Calls per second = 321,445    strIn = "24-08-2021 15:18:01"      DateOrder = 1  Result as expected? True
'Calls per second = 200,997    strIn = "24-08-2021 15:18:01.123"  DateOrder = 1  Result as expected? True
'Calls per second = 375,057    strIn = "2021-08-24 15:18:01"      DateOrder = 2  Result as expected? True
'Calls per second = 207,064    strIn = "2021-08-24 15:18:01.123"  DateOrder = 2  Result as expected? True
'Calls per second = 475,397    strIn = "2021-08-24 15:18:01.123x" DateOrder = 2  Result as expected? True
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastToDate()

          Const N As Long = 1000000
          Dim Converted As Boolean
          Dim DateOrder As Long
          Dim DateSeparator As String
          Dim dtOut As Date
          Dim Expected As Date
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim strIn As String
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
          Dim t1 As Double
          Dim t2 As Double

          '0 = month-day-year, 1 = day-month-year, 2 = year-month-day
1         SysDateOrder = Application.International(xlDateOrder)
2         SysDateSeparator = Application.International(xlDateSeparator)

3         Debug.Print String(100, "=")
4         Debug.Print "Running SpeedTest_CastToDate " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
5         Debug.Print "SysDateOrder = " & SysDateOrder
6         Debug.Print "SysDateSeparator = " & SysDateSeparator
7         Debug.Print "N = " & Format$(N, "###,###")
          
8         For k = 1 To 12
9             For j = 1 To 1 'Maybe do multiple times to test for variability or results.
10                dtOut = 0
11                Converted = False
12                Select Case k
                      Case 1
13                        DateOrder = 2
14                        DateSeparator = "-"
15                        strIn = "foo" 'Contains no date separator, so rejected quickly by CastToDate
16                        Expected = CDate(0)
17                    Case 2
18                        DateOrder = 2
19                        DateSeparator = "-"
20                        strIn = "foo-bar" 'Contains only one date separator, so rejected quickly by CastToDate
21                        Expected = CDate(0)
22                    Case 3
23                        DateOrder = 0 'month-day-year
24                        DateSeparator = "-"
25                        strIn = "09-07-2021"
26                        Expected = CDate("2021-Sep-07")
27                    Case 4
28                        DateOrder = 1 'day-month-year
29                        DateSeparator = "-"
30                        strIn = "07-09-2021"
31                        Expected = CDate("2021-Sep-07")
32                    Case 5
33                        DateOrder = 2   'year-month-day
34                        DateSeparator = "-"
35                        strIn = "2021-09-07"
36                        Expected = CDate("2021-Sep-07")
37                    Case 6
38                        DateOrder = 0
39                        DateSeparator = "-"
40                        strIn = "08-24-2021 15:18:01" 'date with time, no fractions of second
41                        Expected = CDate("2021-Aug-24 15:18:01")
42                    Case 7
43                        DateOrder = 0
44                        DateSeparator = "-"
45                        strIn = "08-24-2021 15:18:01.123" 'date with time, with fractions of second
46                        Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
47                    Case 8
48                        DateOrder = 1
49                        DateSeparator = "-"
50                        strIn = "24-08-2021 15:18:01" 'date with time, no fractions of second
51                        Expected = CDate("2021-Aug-24 15:18:01")
52                    Case 9
53                        DateOrder = 1
54                        DateSeparator = "-"
55                        strIn = "24-08-2021 15:18:01.123" 'date with time, with fractions of second
56                        Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
57                    Case 10
58                        DateOrder = 2
59                        DateSeparator = "-"
60                        strIn = "2021-08-24 15:18:01" 'date with time, no fractions of second
61                        Expected = CDate("2021-Aug-24 15:18:01")
62                    Case 11
63                        DateOrder = 2
64                        DateSeparator = "-"
65                        strIn = "2021-08-24 15:18:01.123" 'date with time, with fractions of second
66                        Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400
67                    Case 12
68                        DateOrder = 2
69                        DateSeparator = "-"
70                        strIn = "2021-08-24 15:18:01.123x" 'Nearly a date, but final "x" stops it being so
71                        Expected = CDate(0)
72                End Select

73                t1 = ElapsedTime()
74                For i = 1 To N
75                    CastToDate strIn, dtOut, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
76                Next i
77                t2 = ElapsedTime()
                  Dim PrintThis As String
78                PrintThis = "Calls per second = " & Format$(N / (t2 - t1), "###,###")
79                If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
80                PrintThis = PrintThis & "strIn = """ & strIn & """"
81                If Len(PrintThis) < 65 Then PrintThis = PrintThis & String(65 - Len(PrintThis), " ")
82                PrintThis = PrintThis & "DateOrder = " & DateOrder & "  Result as expected? " & (Expected = dtOut)
                  
83                Debug.Print PrintThis
84                DoEvents 'kick Immediate window to life
85            Next j
86        Next k

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate, sub-routine of ConvertField
' Purpose    : In-place conversion of a string that looks like a date into a Long or Date. No error if string cannot be
'              converted to date. Converts Dates, DateTimes and Times. Times in very simple format hh:mm:ss
'              Does not handle ISO8601 - see alternative function CastISO8601
' Parameters :
'  strIn           : String
'  dtOut           : Result of cast
'  DateOrder       : The date order respected by the contents of strIn. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  DateSeparator   : The date separator used by the input
'  SysDateOrder    : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  SysDateSeparator: The Windows system date separator
'  Converted       : Boolean flipped to TRUE if conversion takes place
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDate(ByVal strIn As String, ByRef dtOut As Date, ByVal DateOrder As Long, _
        ByVal DateSeparator As String, ByVal SysDateOrder As Long, ByVal SysDateSeparator As String, _
        ByRef Converted As Boolean)
          
          Dim D As String
          Dim M As String
          Dim pos1 As Long 'First date separator
          Dim pos2 As Long 'Second date separator
          Dim pos3 As Long 'Space to separate date from time
          Dim pos4 As Long 'decimal point for fractions of a second
          Dim Converted2 As Boolean
          Dim HasFractionalSecond As Boolean
          Dim HasTimePart As Boolean
          Dim TimePart As String
          Dim TimePartConverted As Date
          Dim y As String
          
1         On Error GoTo ErrHandler
          
2         pos1 = InStr(strIn, DateSeparator)
3         If pos1 = 0 Then Exit Sub
4         pos2 = InStr(pos1 + 1, strIn, DateSeparator)
5         If pos2 = 0 Then Exit Sub
6         pos3 = InStr(pos2 + 1, strIn, " ")
          
7         HasTimePart = pos3 > 0
          
8         If Not HasTimePart Then
9             If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
10                If pos1 = 5 Then
11                    dtOut = CDate(strIn)
12                    Converted = True
13                    Exit Sub
14                End If
15            ElseIf DateOrder = SysDateOrder Then
16                dtOut = CDate(strIn)
17                Converted = True
18                Exit Sub
19            End If
20            If DateOrder = 0 Then 'M-D-Y
21                M = Left$(strIn, pos1 - 1)
22                D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
23                y = Mid$(strIn, pos2 + 1)
24            ElseIf DateOrder = 1 Then 'D-M-Y
25                D = Left$(strIn, pos1 - 1)
26                M = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
27                y = Mid$(strIn, pos2 + 1)
28            ElseIf DateOrder = 2 Then 'Y-M-D
29                y = Left$(strIn, pos1 - 1)
30                M = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
31                D = Mid$(strIn, pos2 + 1)
32            Else
33                Throw "DateOrder must be 0, 1, or 2"
34            End If
35            If SysDateOrder = 0 Then
36                dtOut = CDate(M & SysDateSeparator & D & SysDateSeparator & y)
37                Converted = True
38            ElseIf SysDateOrder = 1 Then
39                dtOut = CDate(D & SysDateSeparator & M & SysDateSeparator & y)
40                Converted = True
41            ElseIf SysDateOrder = 2 Then
42                dtOut = CDate(y & SysDateSeparator & M & SysDateSeparator & D)
43                Converted = True
44            End If
45            Exit Sub
46        End If

47        pos4 = InStr(pos3 + 1, strIn, ".")
48        HasFractionalSecond = pos4 > 0

49        If DateOrder = 0 Then 'M-D-Y
50            M = Left$(strIn, pos1 - 1)
51            D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
52            y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
53            TimePart = Mid$(strIn, pos3)

54        ElseIf DateOrder = 1 Then 'D-M-Y
55            D = Left$(strIn, pos1 - 1)
56            M = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
57            y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
58            TimePart = Mid$(strIn, pos3)
59        ElseIf DateOrder = 2 Then 'Y-M-D
60            y = Left$(strIn, pos1 - 1)
61            M = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
62            D = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
63            TimePart = Mid$(strIn, pos3)
64        Else
65            Throw "DateOrder must be 0, 1, or 2"
66        End If
67        If Not HasFractionalSecond Then
68            If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
69                If pos1 = 5 Then
70                    dtOut = CDate(strIn)
71                    Converted = True
72                    Exit Sub
73                End If
74            ElseIf DateOrder = SysDateOrder Then
75                dtOut = CDate(strIn)
76                Converted = True
77                Exit Sub
78            End If
          
79            If SysDateOrder = 0 Then
80                dtOut = CDate(M & SysDateSeparator & D & SysDateSeparator & y & TimePart)
81                Converted = True
82            ElseIf SysDateOrder = 1 Then
83                dtOut = CDate(D & SysDateSeparator & M & SysDateSeparator & y & TimePart)
84                Converted = True
85            ElseIf SysDateOrder = 2 Then
86                dtOut = CDate(y & SysDateSeparator & M & SysDateSeparator & D & TimePart)
87                Converted = True
88            End If
89        Else 'CDate does not cope with fractional seconds, so use CastToTimeB
90            CastToTimeB Mid$(TimePart, 2), TimePartConverted, Converted2
91            If Converted2 Then
92                If SysDateOrder = 0 Then
93                    dtOut = CDate(M & SysDateSeparator & D & SysDateSeparator & y) + TimePartConverted
94                    Converted = True
95                ElseIf SysDateOrder = 1 Then
96                    dtOut = CDate(D & SysDateSeparator & M & SysDateSeparator & y) + TimePartConverted
97                    Converted = True
98                ElseIf SysDateOrder = 2 Then
99                    dtOut = CDate(y & SysDateSeparator & M & SysDateSeparator & D) + TimePartConverted
100                   Converted = True
101               End If
102           End If
103       End If

104       Exit Sub
ErrHandler:
          'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NumDimensions(ByVal x As Variant) As Long
          Dim i As Long
          Dim y As Long
1         If Not IsArray(x) Then
2             NumDimensions = 0
3             Exit Function
4         Else
5             On Error GoTo ExitPoint
6             i = 1
7             Do While True
8                 y = LBound(x, i)
9                 i = i + 1
10            Loop
11        End If
ExitPoint:
12        NumDimensions = i - 1
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeSentinels
' Purpose    : Returns a Dictionary keyed on strings for which if a key to the dictionary is a field of the CSV file then
'              that field should be converted to the associated item value. Handles Booleans, Missings and Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub MakeSentinels(ByRef Sentinels As Scripting.Dictionary, ByRef MaxLength As Long, _
        ByRef AnySentinels As Boolean, ByVal ShowBooleansAsBooleans As Boolean, ByVal ShowErrorsAsErrors As Boolean, _
        ByRef ShowMissingsAs As Variant, Optional ByVal TrueStrings As Variant, Optional ByVal FalseStrings As Variant, _
        Optional ByVal MissingStrings As Variant)

          Const Err_FalseStrings As String = "FalseStrings must be omitted or provided as a string or an array of " & _
              "strings that represent Boolean value False"
          Const Err_MissingStrings As String = "MissingStrings must be omitted or provided a string or an array of " & _
              "strings that represent missing values"
          Const Err_ShowMissings As String = "ShowMissingsAs has an illegal value, such as an array or an object"
          Const Err_TrueStrings As String = "TrueStrings must be omitted or provided as string or an array of " & _
              "strings that represent Boolean value True"
          Const Err_TrueStrings2 As String = "TrueStrings has been provided, but type conversion for Booleans is " & _
              "not switched on for any column"
          Const Err_FalseStrings2 As String = "FalseStrings has been provided, but type conversion for Booleans " & _
              "is not switched on for any column"

1         On Error GoTo ErrHandler

2         If IsMissing(ShowMissingsAs) Then ShowMissingsAs = Empty
3         Select Case VarType(ShowMissingsAs)
              Case vbObject, vbArray, vbByte, vbDataObject, vbUserDefinedType, vbVariant
4                 Throw Err_ShowMissings
5         End Select
          
6         If Not IsMissing(MissingStrings) And Not IsEmpty(MissingStrings) Then
7             AddKeysToDict Sentinels, MissingStrings, ShowMissingsAs, Err_MissingStrings
8         End If

9         If ShowBooleansAsBooleans Then
10            If IsMissing(TrueStrings) Or IsEmpty(TrueStrings) Then
11                AddKeysToDict Sentinels, Array("TRUE", "true", "True"), True, Err_TrueStrings
12            Else
13                AddKeysToDict Sentinels, TrueStrings, True, Err_TrueStrings
14            End If
15            If IsMissing(FalseStrings) Or IsEmpty(FalseStrings) Then
16                AddKeysToDict Sentinels, Array("FALSE", "false", "False"), False, Err_FalseStrings
17            Else
18                AddKeysToDict Sentinels, FalseStrings, False, Err_FalseStrings
19            End If
20        Else
21            If Not (IsMissing(TrueStrings) Or IsEmpty(TrueStrings)) Then
22                Throw Err_TrueStrings2
23            End If
24            If Not (IsMissing(FalseStrings) Or IsEmpty(FalseStrings)) Then
25                Throw Err_FalseStrings2
26            End If
27        End If
          
28        If ShowErrorsAsErrors Then
29            AddKeyToDict Sentinels, "#DIV/0!", CVErr(xlErrDiv0)
30            AddKeyToDict Sentinels, "#NAME?", CVErr(xlErrName)
31            AddKeyToDict Sentinels, "#REF!", CVErr(xlErrRef)
32            AddKeyToDict Sentinels, "#NUM!", CVErr(xlErrNum)
33            AddKeyToDict Sentinels, "#NULL!", CVErr(xlErrNull)
34            AddKeyToDict Sentinels, "#N/A", CVErr(xlErrNA)
35            AddKeyToDict Sentinels, "#VALUE!", CVErr(xlErrValue)
36            AddKeyToDict Sentinels, "#SPILL!", CVErr(2045)
37            AddKeyToDict Sentinels, "#BLOCKED!", CVErr(2047)
38            AddKeyToDict Sentinels, "#CONNECT!", CVErr(2046)
39            AddKeyToDict Sentinels, "#UNKNOWN!", CVErr(2048)
40            AddKeyToDict Sentinels, "#GETTING_DATA!", CVErr(2043)
41            AddKeyToDict Sentinels, "#FIELD!", CVErr(2049)
42            AddKeyToDict Sentinels, "#CALC!", CVErr(2050)
43        End If

          Dim k As Variant
44        MaxLength = 0
45        For Each k In Sentinels.Keys
46            If Len(k) > MaxLength Then MaxLength = Len(k)
47        Next
48        AnySentinels = Sentinels.Count > 0

49        Exit Sub
ErrHandler:
50        Throw "#MakeSentinels: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeysToDict, Sub-routine of MakeSentinels
' Purpose    : Broadcast AddKeyToDict over an array of keys.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeysToDict(ByRef Sentinels As Scripting.Dictionary, ByVal Keys As Variant, ByVal item As Variant, _
        ByVal FriendlyErrorString As String)

          Dim i As Long
          Dim j As Long
        
1         On Error GoTo ErrHandler
        
2         If TypeName(Keys) = "Range" Then
3             Keys = Keys.Value
4         End If
          
5         If VarType(Keys) = vbString Then
6             If InStr(Keys, ",") > 0 Then
7                 Keys = VBA.Split(Keys, ",")
8             End If
9         End If
          
10        Select Case NumDimensions(Keys)
              Case 0
11                AddKeyToDict Sentinels, Keys, item, FriendlyErrorString
12            Case 1
13                For i = LBound(Keys) To UBound(Keys)
14                    AddKeyToDict Sentinels, Keys(i), item, FriendlyErrorString
15                Next i
16            Case 2
17                For i = LBound(Keys, 1) To UBound(Keys, 1)
18                    For j = LBound(Keys, 2) To UBound(Keys, 2)
19                        AddKeyToDict Sentinels, Keys(i, j), item, FriendlyErrorString
20                    Next j
21                Next i
22            Case Else
23                Throw FriendlyErrorString
24        End Select
25        Exit Sub
ErrHandler:
26        Throw "#AddKeysToDict: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeyToDict, Sub-routine of MakeSentinels
' Purpose    : Wrap .Add method to have more helpful error message if things go awry.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeyToDict(ByRef Sentinels As Scripting.Dictionary, ByVal Key As Variant, ByVal item As Variant, _
        Optional ByVal FriendlyErrorString As String)

          Dim FoundRepeated As Boolean

1         On Error GoTo ErrHandler

2         If VarType(Key) <> vbString Then Throw FriendlyErrorString & " but '" & CStr(Key) & "' is of type " & TypeName(Key)
3         If Len(Key) = 0 Then Exit Sub
          
4         If Not Sentinels.Exists(Key) Then
5             Sentinels.Add Key, item
6         Else
7             FoundRepeated = True
8             If VarType(item) = VarType(Sentinels.item(Key)) Then
9                 If item = Sentinels.item(Key) Then
10                    FoundRepeated = False
11                End If
12            End If
13        End If

14        If FoundRepeated Then
15            Throw "There is a conflicting definition of what the string '" & Key & _
                  "' should be converted to, both the " & TypeName(item) & " value '" & CStr(item) & _
                  "' and the " & TypeName(Sentinels.item(Key)) & " value '" & CStr(Sentinels.item(Key)) & _
                  "' have been specified. Please check the TrueStrings, FalseStrings and MissingStrings arguments."
16        End If

17        Exit Sub
ErrHandler:
18        Throw "#AddKeyToDict: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_Sentinels
' Purpose    : Test speed of accessing the sentinels dictionary, using similar approach to that employed in method
'              ConvertField.
'
' Results:  On Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM
'
'Running SpeedTest_Sentinels 2021-08-25T15:01:33
'Conversions per second = 90,346,968       Field = "This string is longer than the longest sentinel, which is 14"
'Conversions per second = 20,976,150       Field = "mini" (Not a sentinel, but shorter than the longest sentinel)
'Conversions per second = 9,295,050        Field = "True" (A sentinel, one of the elements of TrueStrings)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_Sentinels()
          
          Const N As Long = 10000000
          Dim AnySentinels As Boolean
          Dim Comment As String
          Dim Field As String
          Dim i As Long
          Dim j As Long
          Dim MaxLength As Long
          Dim Res As Variant
          Dim Sentinels As Scripting.Dictionary
          Dim t1 As Double
          Dim t2 As Double

1         On Error GoTo ErrHandler
          
2         Set Sentinels = New Scripting.Dictionary
3         MakeSentinels Sentinels, MaxLength, AnySentinels, _
              ShowBooleansAsBooleans:=True, _
              ShowErrorsAsErrors:=True, _
              ShowMissingsAs:=Empty, _
              TrueStrings:=Array("True", "T"), _
              FalseStrings:=Array("False", "F"), _
              MissingStrings:=Array("NA", "-999")
          
          Dim Converted As Boolean
          
4         Debug.Print "Running SpeedTest_Sentinels " & Format$(Now(), "yyyy-mm-ddThh:mm:ss")
          
5         For j = 1 To 3
          
6             Select Case j
                  Case 1
7                     Field = "This string is longer than the longest sentinel, which is 14"
8                 Case 2
9                     Field = "mini"
10                    Comment = "Not a sentinel, but shorter than the longest sentinel"
11                Case 3
12                    Field = "True"
13                    Comment = "A sentinel, one of the elements of TrueStrings"
14            End Select

15            t1 = ElapsedTime()
16            For i = 1 To N
17                If Len(Field) <= MaxLength Then
18                    If Sentinels.Exists(Field) Then
19                        Res = Sentinels.item(Field)
20                        Converted = True
21                    End If
22                End If
23            Next i
24            t2 = ElapsedTime()

25            Debug.Print "Conversions per second = " & Format$(N / (t2 - t1), "###,###"), _
                  "Field = """ & Field & """" & IIf(Comment = vbNullString, vbNullString, " (" & Comment & ")")

26        Next j

27        Exit Sub
ErrHandler:
28        MsgBox "#SpeedTest_Sentinels: " & Err.Description & "!"

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseISO8601
' Purpose    : Test harness for calling from spreadsheets
' -----------------------------------------------------------------------------------------------------------------------
Public Function ParseISO8601(ByVal strIn As String) As Variant
          Dim Converted As Boolean
          Dim dtOut As Date

1         On Error GoTo ErrHandler
2         CastISO8601 strIn, dtOut, Converted, True, True

3         If Converted Then
4             ParseISO8601 = dtOut
5         Else
6             ParseISO8601 = "#Not recognised as ISO8601 date!"
7         End If
8         Exit Function
ErrHandler:
9         ParseISO8601 = "#ParseISO8601: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTime
' Purpose    : Cast strings that represent a time to a date, no handling of TimeZone.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTime(ByVal strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)

1         On Error GoTo ErrHandler
          
2         dtOut = CDate(strIn)
3         If dtOut <= 1 Then
4             Converted = True
5         End If
          
6         Exit Sub
ErrHandler:
          'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTimeB
' Purpose    : CDate does not correctly cope with times such as '04:20:10.123 am' or '04:20:10.123', i.e, times with a
'              fractional second, so this method is called after CastToTime
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTimeB(ByVal strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)
          Static rx As VBScript_RegExp_55.RegExp
          Dim DecPointAt As Long
          Dim FractionalSecond As Double
          Dim SpaceAt As Long
          
1         On Error GoTo ErrHandler
2         If rx Is Nothing Then
3             Set rx = New RegExp
4             With rx
5                 .IgnoreCase = True
6                 .Pattern = "^[0-2]?[0-9]:[0-5]?[0-9]:[0-5]?[0-9](\.[0-9]+)( am| pm)?$"
7                 .Global = False        'Find first match only
8             End With
9         End If

10        If Not rx.Test(strIn) Then Exit Sub
11        DecPointAt = InStr(strIn, ".")
12        If DecPointAt = 0 Then Exit Sub ' should never happen
13        SpaceAt = InStr(strIn, " ")
14        If SpaceAt = 0 Then SpaceAt = Len(strIn) + 1
15        FractionalSecond = CDbl(Mid$(strIn, DecPointAt, SpaceAt - DecPointAt)) / 86400
          
16        dtOut = CDate(Left$(strIn, DecPointAt - 1) + Mid$(strIn, SpaceAt)) + FractionalSecond
17        Converted = True
18        Exit Sub
ErrHandler:
          'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastISO8601
' Purpose    : Convert ISO8601 formatted datestrings to UTC date. https://xkcd.com/1179/

'Always accepts dates without time
'Format                        Example
'yyyy-mm-dd                    2021-08-23

'If AcceptWithoutTimeZone is True:
'yyyy-mm-ddThh:mm:ss           2021-08-23T08:47:21
'yyyy-mm-ddThh:mm:ss.000       2021-08-23T08:47:20.920

'If AcceptWithTimeZone is True:
'yyyy-mm-ddThh:mm:ssZ          2021-08-23T08:47:21Z
'yyyy-mm-ddThh:mm:ss.000Z      2021-08-23T08:47:20.920Z
'yyyy-mm-ddThh:mm:ss+hh:mm     2021-08-23T08:47:21+05:00
'yyyy-mm-ddThh:mm:ss.000+hh:mm 2021-08-23T08:47:20.920+05:00

' Parameters :
'  StrIn                : The string to be converted
'  DtOut                : The date that the string converts to.
'  Converted            : Did the function convert (true) or reject as not a correctly formatted date (false)
'  AcceptWithoutTimeZone: Should the function accept datetime without time zone given?
'  AcceptWithTimeZone   : Should the function accept datetime with time zone given?

'       IMPORTANT:       WHEN TIMEZONE IS GIVEN THE FUNCTION RETURNS THE TIME IN UTC

' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastISO8601(ByVal strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean, _
        ByVal AcceptWithoutTimeZone As Boolean, ByVal AcceptWithTimeZone As Boolean)

          Dim L As Long
          Dim LocalTime As Double
          Dim MilliPart As Double
          Dim MinusPos As Long
          Dim PlusPos As Long
          Dim Sign As Long
          Dim ZAtEnd As Boolean
          
          Static rxNoNo As VBScript_RegExp_55.RegExp
          Static RxYesNo As VBScript_RegExp_55.RegExp
          Static RxNoYes As VBScript_RegExp_55.RegExp
          Static rxYesYes As VBScript_RegExp_55.RegExp
          Static rxExists As Boolean

1         On Error GoTo ErrHandler
          
2         If Not rxExists Then
3             Set rxNoNo = New RegExp
              'Reject datetime
4             With rxNoNo
5                 .IgnoreCase = False
6                 .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9]$"
7                 .Global = False
8             End With
              
              'Accept datetime without time zone, reject datetime with timezone
9             Set RxYesNo = New RegExp
10            With RxYesNo
11                .IgnoreCase = False
12                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?)?$"
13                .Global = False
14            End With
              
              'Reject datetime without time zone, accept datetime with timezone
15            Set RxNoYes = New RegExp
16            With RxNoYes
17                .IgnoreCase = False
18                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?(Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?$"
19                .Global = False
20            End With
              
              'Accept datetime, both with and without timezone
21            Set rxYesYes = New RegExp
22            With rxYesYes
23                .IgnoreCase = False
24                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?((Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?)?$"
25                .Global = False
26            End With
27            rxExists = True
28        End If
          
29        L = Len(strIn)

30        If L = 10 Then
31            If rxNoNo.Test(strIn) Then
                  'This works irrespective of Windows regional settings
32                dtOut = CDate(strIn)
33                Converted = True
34                Exit Sub
35            End If
36        ElseIf L < 10 Then
37            Converted = False
38            Exit Sub
39        ElseIf L > 40 Then
40            Converted = False
41            Exit Sub
42        End If

43        Converted = False
          
44        If AcceptWithoutTimeZone Then
45            If AcceptWithTimeZone Then
46                If Not rxYesYes.Test(strIn) Then Exit Sub
47            Else
48                If Not RxYesNo.Test(strIn) Then Exit Sub
49            End If
50        Else
51            If AcceptWithTimeZone Then
52                If Not RxNoYes.Test(strIn) Then Exit Sub
53            Else
54                If Not rxNoNo.Test(strIn) Then Exit Sub
55            End If
56        End If
          
          'Replace the "T" separator
57        Mid$(strIn, 11, 1) = " "
          
58        If L = 19 Then
59            dtOut = CDate(strIn)
60            Converted = True
61            Exit Sub
62        End If

63        If Right$(strIn, 1) = "Z" Then
64            Sign = 0
65            ZAtEnd = True
66        Else
67            PlusPos = InStr(20, strIn, "+")
68            If PlusPos > 0 Then
69                Sign = 1
70            Else
71                MinusPos = InStr(20, strIn, "-")
72                If MinusPos > 0 Then
73                    Sign = -1
74                End If
75            End If
76        End If

77        If Mid$(strIn, 20, 1) = "." Then 'Have fraction of a second
78            Select Case Sign
                  Case 0
                      'Example: "2021-08-23T08:47:20.920Z"
79                    MilliPart = CDbl(Mid$(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
80                Case 1
                      'Example: "2021-08-23T08:47:20.920+05:00"
81                    MilliPart = CDbl(Mid$(strIn, 20, PlusPos - 20))
82                Case -1
                      'Example: "2021-08-23T08:47:20.920-05:00"
83                    MilliPart = CDbl(Mid$(strIn, 20, MinusPos - 20))
84            End Select
85        End If
          
86        LocalTime = CDate(Left$(strIn, 19)) + MilliPart / 86400

          Dim Adjust As Date
87        Select Case Sign
              Case 0
88                dtOut = LocalTime
89                Converted = True
90                Exit Sub
91            Case 1
92                If L <> PlusPos + 5 Then Exit Sub
93                Adjust = CDate(Right$(strIn, 5))
94                dtOut = LocalTime - Adjust
95                Converted = True
96            Case -1
97                If L <> MinusPos + 5 Then Exit Sub
98                Adjust = CDate(Right$(strIn, 5))
99                dtOut = LocalTime + Adjust
100               Converted = True
101       End Select

102       Exit Sub
ErrHandler:
          'Was not recognised as ISO8601 date
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTest_CastISO8601
' Purpose    : Testing speed of CastISO8601

'Example output: (Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM)

'Running SpeedTest_CastISO8601 2021-Sep-07 22:03:27
'N = 5,000,000
'Calls per second = 1,052,769  strIn = "xxxxxxxxxxxxxxxxxxxxxxxxxxx..."  Result as expected? True
'Calls per second = 2,436,414  strIn = "Foo"                             Result as expected? True
'Calls per second = 1,718,279  strIn = "xxxxxxxxxxxx"                    Result as expected? True
'Calls per second = 1,718,023  strIn = "xxxx-xxxxxxx"                    Result as expected? True
'Calls per second = 587,754    strIn = "2021-08-24T15:18:01.123+05:0x"   Result as expected? True
'Calls per second = 574,610    strIn = "2021-08-23"                      Result as expected? True
'Calls per second = 348,325    strIn = "2021-08-24T15:18:01"             Result as expected? True
'Calls per second = 247,093    strIn = "2021-08-23T08:47:21.123"         Result as expected? True
'Calls per second = 221,942    strIn = "2021-08-24T15:18:01+05:00"       Result as expected? True
'Calls per second = 191,331    strIn = "2021-08-24T15:18:01.123+05:00"   Result as expected? True
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTest_CastISO8601()

          Const N As Long = 5000000
          Dim Converted As Boolean
          Dim dtOut As Date
          Dim Expected As Date
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim PrintThis As String
          Dim strIn As String
          Dim t1 As Double
          Dim t2 As Double

1         Debug.Print String(100, "=")
2         Debug.Print "Running SpeedTest_CastISO8601 " & Format$(Now(), "yyyy-mmm-dd hh:mm:ss")
3         Debug.Print "N = " & Format$(N, "###,###")
4         For k = 0 To 9
5             For j = 1 To 1
6                 dtOut = 0
7                 Select Case k
                      Case 0
8                         strIn = String(10000, "x")
9                         Expected = CDate(0)
10                    Case 1
11                        strIn = "Foo" ' less than 10 in length
12                        Expected = CDate(0)
13                    Case 2
14                        strIn = "xxxxxxxxxxxx" '5th character not "-"
15                        Expected = CDate(0)
16                    Case 3
17                        strIn = "xxxx-xxxxxxx" 'rejected by RegEx
18                        Expected = CDate(0)
19                    Case 4
20                        strIn = "2021-08-24T15:18:01.123+05:0x" ' rejected by regex
21                        Expected = CDate(0)
22                    Case 5
23                        strIn = "2021-08-23"
24                        Expected = CDate("2021-Aug-23")
25                    Case 6
26                        strIn = "2021-08-24T15:18:01"
27                        Expected = CDate("2021-Aug-24 15:18:01")
28                    Case 7
29                        strIn = "2021-08-23T08:47:21.123"
30                        Expected = CDate("2021-Aug-23 08:47:21") + 0.123 / 86400
31                    Case 8
32                        strIn = "2021-08-24T15:18:01+05:00"
33                        Expected = CDate("2021-Aug-24 15:18:01") - 5 / 24
34                    Case 9
35                        strIn = "2021-08-24T15:18:01.123+05:00"
36                        Expected = CDate("2021-Aug-24 15:18:01") + 0.123 / 86400 - 5 / 24
37                End Select

38                t1 = ElapsedTime()
39                For i = 1 To N
40                    CastISO8601 strIn, dtOut, Converted, True, True
41                Next i
42                t2 = ElapsedTime()
                  
43                PrintThis = "Calls per second = " & Format$(N / (t2 - t1), "###,###")
44                If Len(PrintThis) < 30 Then PrintThis = PrintThis & String(30 - Len(PrintThis), " ")
45                If Len(strIn) > 30 Then
46                    PrintThis = PrintThis & "strIn = """ & Left$(strIn, 27) & "..."""
47                Else
48                    PrintThis = PrintThis & "strIn = """ & strIn & """"
49                End If
50                If Len(PrintThis) < 70 Then PrintThis = PrintThis & String(70 - Len(PrintThis), " ")
51                PrintThis = PrintThis & "  Result as expected? " & (Expected = dtOut)
                  
52                Debug.Print PrintThis
53                DoEvents 'kick Immediate window to life
54            Next j
55        Next k

56    End Sub

      'See "gogeek"'s post at _
 https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLocalOffsetToUTC
' Purpose    : Get the PC's offset to UTC.
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLocalOffsetToUTC() As Double
          Dim dt As Object
          Dim TimeNow As Date
          Dim UTC As Date
1         On Error GoTo ErrHandler
2         TimeNow = Now()

3         Set dt = CreateObject("WbemScripting.SWbemDateTime")
4         dt.SetVarDate TimeNow
5         UTC = dt.GetVarDate(False)
6         GetLocalOffsetToUTC = (TimeNow - UTC)

7         Exit Function
ErrHandler:
8         Throw "#GetLocalOffsetToUTC: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISOZFormatString
' Purpose    : Returns the format string required to save datetimes with timezone under the assumton that the offset to
'              UTC is the same as the curent offset on this PC - use with care, Daylight saving may mean that that's not
'              a correct assumption for all the dates in a set of data...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISOZFormatString() As String
          Dim RightChars As String
          Dim TimeZone As String

1         On Error GoTo ErrHandler
2         TimeZone = GetLocalOffsetToUTC()

3         If TimeZone = 0 Then
4             RightChars = "Z"
5         ElseIf TimeZone > 0 Then
6             RightChars = "+" & Format$(TimeZone, "hh:mm")
7         Else
8             RightChars = "-" & Format$(Abs(TimeZone), "hh:mm")
9         End If
10        ISOZFormatString = "yyyy-mm-ddT:hh:mm:ss" & RightChars

11        Exit Function
ErrHandler:
12        Throw "#ISOZFormatString: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseTextFile
' Purpose    : Convert a text file to a 2-dim array with one column, one line of file to one element of array, works for
'              files with any style of line endings - Windows, Mac, Unix, or a mixture of line endings.
' Parameters :
'  FileNameOrContents : FileName or CSV-style string.
'  isFile             : If True then fist argument is the name of a file, else it's a CSV-style string.
'  useADODB           : Should the file be read using ADODB.Stream rather than Scripting.TextStream? Necessary only for
'                       UTF-8 files.
'  CharSet            : Used only if useADODB is True and isFile is True.
'  TriState           : Used only if useADODB is False and isFile is True.
'  SkipToLine         : Return starts at this line of the file.
'  NumLinesToReturn   : This many lines are returned. Pass zero for all lines from SkipToLine.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseTextFile(FileNameOrContents As String, ByVal isFile As Boolean, ByVal useADODB As Boolean, _
        ByVal CharSet As String, ByVal TriState As Long, ByVal SkipToLine As Long, ByVal NumLinesToReturn As Long, _
        ByVal CallingFromWorksheet As Boolean) As Variant

          Const Err_FileEmpty As String = "File is empty"
          Dim Buffer As String
          Dim BufferUpdatedTo As Long
          Dim FoundCR As Boolean
          Dim HaveReachedSkipToLine As Boolean
          Dim i As Long 'Index to read from Buffer
          Dim j As Long 'Index to write to Starts, Lengths
          Dim Lengths() As Long
          Dim NumLinesFound As Long
          Dim PosCR As Long
          Dim PosLF As Long
          Dim ReturnArray() As String
          Dim SearchFor() As String
          Dim Starts() As Long
          Dim Stream As Object
          Dim Streaming As Boolean
          Dim tmp As Long
          Dim Which As Long
          Dim MSLIA As Long
          Dim Err_StringTooLong As String

1         On Error GoTo ErrHandler
          
2         If isFile Then
3             If useADODB Then
4                 Set Stream = CreateObject("ADODB.Stream")
5                 Stream.CharSet = CharSet
6                 Stream.Open
7                 Stream.LoadFromFile FileNameOrContents
8                 If Stream.EOS Then Throw Err_FileEmpty
9             Else
10                If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
11                Set Stream = m_FSO.GetFile(FileNameOrContents).OpenAsTextStream(ForReading, TriState)
12                If Stream.atEndOfStream Then Throw Err_FileEmpty
13            End If
          
14            If NumLinesToReturn = 0 Then
15                Buffer = ReadAllFromStream(Stream)
16                Streaming = False
17            Else
18                GetMoreFromStream Stream, vbNullString, vbNullString, Buffer, BufferUpdatedTo
19                Streaming = True
20            End If
21        Else
22            Buffer = FileNameOrContents
23            Streaming = False
24        End If
             
25        If Streaming Then
26            ReDim SearchFor(1 To 2)
27            SearchFor(1) = vbLf
28            SearchFor(2) = vbCr
29        End If

30        ReDim Starts(1 To 8): ReDim Lengths(1 To 8)
          
31        If Not Streaming Then
              'Ensure Buffer terminates with vbCrLf
32            If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
33                Buffer = Buffer & vbCrLf
34            ElseIf Right$(Buffer, 1) = vbCr Then
35                Buffer = Buffer & vbLf
36            End If
37            BufferUpdatedTo = Len(Buffer)
38        End If
          
39        NumLinesFound = 0
40        i = 0: j = 1
          
41        Starts(1) = i + 1
42        If SkipToLine = 1 Then HaveReachedSkipToLine = True

43        Do
44            If Not Streaming Then
45                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
46                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
47                If PosCR < PosLF Then
48                    FoundCR = True
49                    i = PosCR
50                Else
51                    FoundCR = False
52                    i = PosLF
53                End If
54            Else
55                i = SearchInBuffer(SearchFor, i + 1, Stream, useADODB, vbNullString, _
                      vbNullString, Which, Buffer, BufferUpdatedTo)
56                FoundCR = (Which = 2)
57            End If

58            If i >= BufferUpdatedTo + 1 Then
59                Exit Do
60            End If

61            If j + 1 > UBound(Starts) Then
62                ReDim Preserve Starts(1 To UBound(Starts) * 2)
63                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
64            End If

65            Lengths(j) = i - Starts(j)
66            If FoundCR Then
67                If Mid$(Buffer, i + 1, 1) = vbLf Then
                      'Ending is Windows rather than Mac or Unix.
68                    i = i + 1
69                End If
70            End If
                          
71            Starts(j + 1) = i + 1
                          
72            j = j + 1
73            NumLinesFound = NumLinesFound + 1
74            If Not HaveReachedSkipToLine Then
75                If NumLinesFound = SkipToLine - 1 Then
76                    HaveReachedSkipToLine = True
77                    tmp = Starts(j)
78                    ReDim Starts(1 To 8): ReDim Lengths(1 To 8)
79                    j = 1: NumLinesFound = 0
80                    Starts(1) = tmp
81                End If
82            ElseIf NumLinesToReturn > 0 Then
83                If NumLinesFound = NumLinesToReturn Then
84                    Exit Do
85                End If
86            End If
87        Loop
         
88        If SkipToLine > NumLinesFound Then
89            If NumLinesToReturn = 0 Then 'Attempting to read from SkipToLine to the end of the file, but that would _
                                            be zero or a negative number of rows. So throw an error.
                                   
90                Throw "SkipToLine (" & CStr(SkipToLine) & ") exceeds the number of lines in the file (" & _
                      CStr(NumLinesFound) & ")"
91            Else
                  'Attempting to read a set number of rows, function will return an array of null strings
92                NumLinesFound = 0
93            End If
94        End If
95        If NumLinesToReturn = 0 Then NumLinesToReturn = NumLinesFound

96        ReDim ReturnArray(1 To NumLinesToReturn, 1 To 1)
97        MSLIA = MaxStringLengthInArray()
98        For i = 1 To MinLngs(NumLinesToReturn, NumLinesFound)
99            If CallingFromWorksheet Then
100               If Lengths(i) > MSLIA Then
101                   Err_StringTooLong = "Line " & Format$(i, "#,###") & " of the file is of length " + Format$(Lengths(i), "###,###")
102                   If MSLIA >= 32767 Then
103                       Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " + Format$(MSLIA, "####,####")
104                   Else
105                       Err_StringTooLong = Err_StringTooLong & _
                              ". An array containing a string longer than " + Format$(MSLIA, "###,###") + _
                              " cannot be returned from VBA to an Excel worksheet"
106                   End If
107                   Throw Err_StringTooLong
108               End If
109           End If
110           ReturnArray(i, 1) = Mid$(Buffer, Starts(i), Lengths(i))
111       Next i

112       ParseTextFile = ReturnArray

113       Exit Function
ErrHandler:
114       Throw "#ParseTextFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Private Function ElapsedTime() As Double
          Dim a As Currency
          Dim b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         QueryPerformanceFrequency b
4         ElapsedTime = a / b

5         Exit Function
ErrHandler:
6         Throw "#ElapsedTime: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if Excel's Function Wizard is active to allow early exit in slow functions.
' https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
          
1         On Error GoTo ErrHandler
2         If Not Application.CommandBars.item("Standard").Controls.item(1).Enabled Then
3             FunctionWizardActive = True
4         End If

5         Exit Function
ErrHandler:
6         Throw "#FunctionWizardActive: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal ErrorString As String)
1         Err.Raise vbObjectError + 1, , ErrorString
End Sub

'' -----------------------------------------------------------------------------------------------------------------------
'' Procedure : ThrowIfError
'' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
''             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
''             methods to be used from VBA code while keeping error handling robust
''             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
'' -----------------------------------------------------------------------------------------------------------------------
'Public Function ThrowIfError(ByRef Data As Variant) As Variant
'    ThrowIfError = Data
'    If VarType(Data) = vbString Then
'        If Left$(Data, 1) = "#" Then
'            If Right$(Data, 1) = "!" Then
'                Throw CStr(Data)
'            End If
'        End If
'    End If
'End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Purpose    : Returns True if FileName exists on disk, False o.w.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileExists(ByVal FileName As String) As Boolean
          Dim F As Scripting.file
1         On Error GoTo ErrHandler
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         Set F = m_FSO.GetFile(FileName)
4         FileExists = True
5         Exit Function
ErrHandler:
6         FileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FolderExists(ByVal FolderPath As String) As Boolean
          Dim F As Scripting.Folder
          
1         On Error GoTo ErrHandler
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
          
3         Set F = m_FSO.GetFolder(FolderPath)
4         FolderExists = True
5         Exit Function
ErrHandler:
6         FolderExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileDelete
' Purpose    : Delete a file, returns True or error string.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FileDelete(ByVal FileName As String)
          Dim F As Scripting.file
1         On Error GoTo ErrHandler

2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         Set F = m_FSO.GetFile(FileName)
4         F.Delete

5         Exit Sub
ErrHandler:
6         Throw "#FileDelete: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful throws an error.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CreatePath(ByVal FolderPath As String)

          Dim F As Scripting.Folder
          Dim i As Long
          Dim ParentFolderName As String
          Dim ThisFolderName As String

1         On Error GoTo ErrHandler

2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

3         If Left$(FolderPath, 2) = "\\" Then
4         ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or _
              Asc(UCase$(Left$(FolderPath, 1))) < 65 Or _
              Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
5             Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
                  "UNC folder name"
6         End If

7         FolderPath = Replace(FolderPath, "/", "\")

8         If Right$(FolderPath, 1) <> "\" Then
9             FolderPath = FolderPath & "\"
10        End If

11        If FolderExists(FolderPath) Then
12            GoTo EarlyExit
13        End If

          'Go back until we find parent folder that does exist
14        For i = Len(FolderPath) - 1 To 3 Step -1
15            If Mid$(FolderPath, i, 1) = "\" Then
16                If FolderExists(Left$(FolderPath, i)) Then
17                    Set F = m_FSO.GetFolder(Left$(FolderPath, i))
18                    ParentFolderName = Left$(FolderPath, i)
19                    Exit For
20                End If
21            End If
22        Next i

23        If F Is Nothing Then Throw "Cannot create folder " & Left$(FolderPath, 3)

          'now add folders one level at a time
24        For i = Len(ParentFolderName) + 1 To Len(FolderPath)
25            If Mid$(FolderPath, i, 1) = "\" Then
                  
26                ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, _
                      i - 1 - InStrRev(FolderPath, "\", i - 1))
27                F.SubFolders.Add ThisFolderName
28                Set F = m_FSO.GetFolder(Left$(FolderPath, i))
29            End If
30        Next i

EarlyExit:
31        Set F = m_FSO.GetFolder(FolderPath)
32        Set F = Nothing

33        Exit Sub
ErrHandler:
34        Throw "#CreatePath: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromPath
' Purpose    : Split file-with-path to file name (if ReturnFileName is True) or path otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileFromPath(ByVal FullFileName As String, Optional ByVal ReturnFileName As Boolean = True) As Variant
          Dim SlashPos As Long
          Dim SlashPos2 As Long

1         On Error GoTo ErrHandler

2         SlashPos = InStrRev(FullFileName, "\")
3         SlashPos2 = InStrRev(FullFileName, "/")
4         If SlashPos2 > SlashPos Then SlashPos = SlashPos2
5         If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

6         If ReturnFileName Then
7             FileFromPath = Mid$(FullFileName, SlashPos + 1)
8         Else
9             FileFromPath = Left$(FullFileName, SlashPos - 1)
10        End If

11        Exit Function
ErrHandler:
12        Throw "#FileFromPath: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumber
' Purpose   : Is a singleton a number?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsNumber(ByVal x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
2                 IsNumber = True
3         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NCols(Optional TheArray As Variant) As Long
1         If TypeName(TheArray) = "Range" Then
2             NCols = TheArray.Columns.Count
3         ElseIf IsMissing(TheArray) Then
4             NCols = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             NCols = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     NCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
10                Case Else
11                    NCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NRows(Optional TheArray As Variant) As Long
1         If TypeName(TheArray) = "Range" Then
2             NRows = TheArray.Rows.Count
3         ElseIf IsMissing(TheArray) Then
4             NRows = 0
5         ElseIf VarType(TheArray) < vbArray Then
6             NRows = 1
7         Else
8             Select Case NumDimensions(TheArray)
                  Case 1
9                     NRows = 1
10                Case Else
11                    NRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
12            End Select
13        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             Return is always 1-based, even when input is zero-based.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Transpose(ByVal TheArray As Variant) As Variant
          Dim Co As Long
          Dim i As Long
          Dim j As Long
          Dim M As Long
          Dim N As Long
          Dim Result As Variant
          Dim Ro As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, M
3         Ro = LBound(TheArray, 1) - 1
4         Co = LBound(TheArray, 2) - 1
5         ReDim Result(1 To M, 1 To N)
6         For i = 1 To N
7             For j = 1 To M
8                 Result(j, i) = TheArray(i + Ro, j + Co)
9             Next j
10        Next i
11        Transpose = Result
12        Exit Function
ErrHandler:
13        Throw "#Transpose: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
1         If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
2         Force2DArray RangeOrArray, NR, NC
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
          Dim TwoDArray As Variant

1         On Error GoTo ErrHandler

2         Select Case NumDimensions(TheArray)
              Case 0
3                 ReDim TwoDArray(1 To 1, 1 To 1)
4                 TwoDArray(1, 1) = TheArray
5                 TheArray = TwoDArray
6                 NR = 1: NC = 1
7             Case 1
                  Dim i As Long
                  Dim LB As Long
8                 LB = LBound(TheArray, 1)
9                 NR = 1: NC = UBound(TheArray, 1) - LB + 1
10                ReDim TwoDArray(1 To 1, 1 To NC)
11                For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
12                    TwoDArray(1, i) = TheArray(LB + i - 1)
13                Next i
14                TheArray = TwoDArray
15            Case 2
16                NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
17                NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
                  'Nothing to do
18            Case Else
19                Throw "Cannot convert array of dimension greater than two"
20        End Select

21        Exit Sub
ErrHandler:
22        Throw "#Force2DArray: " & Err.Description & "!"
End Sub

Private Function MaxLngs(ByVal x As Long, ByVal y As Long) As Long
1         If x > y Then
2             MaxLngs = x
3         Else
4             MaxLngs = y
5         End If
End Function

Private Function MinLngs(ByVal x As Long, ByVal y As Long) As Long
1         If x > y Then
2             MinLngs = y
3         Else
4             MinLngs = x
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function sCSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVWrite()
          Const Description As String = "Creates a comma-separated file on disk containing Data. Any existing file of " & _
              "the same name is overwritten. If successful, the function returns FileName, " & _
              "otherwise an ""error string"" (starts with `#`, ends with `!`) describing what " & _
              "went wrong."
          Dim ArgDescs() As String

1         On Error GoTo ErrHandler

2         ReDim ArgDescs(1 To 8)
3         ArgDescs(1) = "An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, " & _
              "Excel errors or null values. Data typically has two dimensions, but if Data has only one " & _
              "dimension then the output file has a single column, one field per row."
4         ArgDescs(2) = "The full name of the file, including the path. Alternatively, if FileName is omitted, then the " & _
              "function returns Data converted CSV-style to a string."
5         ArgDescs(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If " & _
              "FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. " & _
              "Double quotes are always escaped by another double quote."
6         ArgDescs(4) = "A format string that determines how dates, including cells formatted as dates, appear in the " & _
              "file. If omitted, defaults to `yyyy-mm-dd`."
7         ArgDescs(5) = "Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use `ISOZ` for " & _
              "ISO8601 format with time zone the same as the PC's clock. Use with care, daylight saving may be " & _
              "inconsistent across the datetimes in data."
8         ArgDescs(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one " & _
              "character."
9         ArgDescs(7) = "Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if this " & _
              "argument is `ANSI` but Data contains characters that cannot be written to an ANSI file. `UTF-8` " & _
              "and `UTF-16` files are written with a byte option mark."
10        ArgDescs(8) = "Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or `Mac`. " & _
              "Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or " & _
              "the strings `CRLF`, `LF` or `CR`."
11        Application.MacroOptions "sCSVWrite", Description, , , , , , , , , ArgDescs
12        Exit Sub

ErrHandler:
13        Debug.Print "Warning: Registration of function sCSVWrite failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same name
'             is overwritten. If successful, the function returns FileName, otherwise an
'             "error string" (starts with `#`, ends with `!`) describing what went wrong.
' Arguments
' Data      : An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans,
'             empty, Excel errors or null values.
' FileName  : The full name of the file, including the path. Alternatively, if FileName is omitted, then
'             the function returns Data converted CSV-style to a string.
' QuoteAllStrings: If TRUE (the default) then all strings in Data are quoted before being written to file. If
'             FALSE only strings containing Delimiter, line feed, carriage return or double
'             quote are quoted. Double quotes are always escaped by another double quote.
' DateFormat: A format string that determines how dates, including cells formatted as dates, appear in
'             the file. If omitted, defaults to `yyyy-mm-dd`.
' DateTimeFormat: Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use
'             `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with
'             care, daylight saving may be inconsistent across the datetimes in data.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Encoding  : Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if
'             this argument is `ANSI` but Data contains characters that cannot be written
'             to an ANSI file. `UTF-8` and `UTF-16` files are written with a byte option
'             mark.
' EOL       : Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or
'             `Mac`. Also supports the line-ending characters themselves (ascii 13 + ascii
'             10, ascii 10, ascii 13) or the strings `CRLF`, `LF` or `CR`.
' TrueString: Sets the text that appears in the file to represent the Boolean value True. Optional,
'             defaulting to "True".
' FalseString: Sets the text that appears in the file to represent the Boolean value False. Optional,
'             defaulting to "False".
'
' Notes     : See also companion function sCSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------------
Public Function sCSVWrite(ByVal Data As Variant, Optional ByVal FileName As String, _
          Optional ByVal QuoteAllStrings As Boolean = True, Optional ByVal DateFormat As String = "YYYY-MM-DD", _
          Optional ByVal DateTimeFormat As String = "ISO", Optional ByVal Delimiter As String = ",", _
          Optional ByVal Encoding As String = "ANSI", Optional ByVal EOL As String = vbNullString, _
          Optional TrueString As String = "True", Optional FalseString As String = "False") As String
Attribute sCSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with `#`, ends with `!`) describing what went wrong."
Attribute sCSVWrite.VB_ProcData.VB_Invoke_Func = " \n26"

          Const DQ As String = """"
          Const Err_Delimiter As String = "Delimiter must have at least one character and cannot start with a " & _
              "double  quote, line feed or carriage return"
          Const Err_Dimensions As String = "Data has more than two dimensions, which is not supported"
          Const Err_Encoding As String = "Encoding must be ""ANSI"" (the default) or ""UTF-8"" or ""UTF-16"""
          
          Dim EOLIsWindows As Boolean
          Dim ErrRet As String
          Dim i As Long
          Dim j As Long
          Dim Lines() As String
          Dim OneLine() As String
          Dim OneLineJoined As String
          Dim Stream As Object
          Dim WriteToFile As Boolean
          Dim Unicode As Boolean

1         On Error GoTo ErrHandler
          
2         Select Case UCase$(Encoding)
              Case "ANSI", "UTF-8", "UTF-16"
3             Case Else
4                 Throw Err_Encoding
5         End Select
6         Select Case TrueString
              Case "False", "false,""FALSE"
7                 Throw "TrueString cannot take the value '" & TrueString & "'"
8         End Select
9         Select Case FalseString
              Case "True", "true", "TRUE"
10                Throw "FalseString cannot take the value '" & FalseString & "'"
11        End Select

12        WriteToFile = Len(FileName) > 0

13        If EOL = vbNullString Then
14            If WriteToFile Then
15                EOL = vbCrLf
16            Else
17                EOL = vbLf
18            End If
19        End If

20        EOL = OStoEOL(EOL, "EOL")
21        EOLIsWindows = EOL = vbCrLf

22        If Len(Delimiter) = 0 Or Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
23            Throw Err_Delimiter
24        End If
          
25        Select Case UCase$(DateTimeFormat)
              Case "ISO"
26                DateTimeFormat = "yyyy-mm-ddThh:mm:ss"
27            Case "ISOZ"
28                DateTimeFormat = ISOZFormatString()
29        End Select

30        If TypeName(Data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
31            Data = Data.Value
32        End If
33        Select Case NumDimensions(Data)
              Case 0
                  Dim tmp() As Variant
34                ReDim tmp(1 To 1, 1 To 1)
35                tmp(1, 1) = Data
36                Data = tmp
37            Case 1
38                ReDim tmp(LBound(Data) To UBound(Data), 1 To 1)
39                For i = LBound(Data) To UBound(Data)
40                    tmp(i, 1) = Data(i)
41                Next i
42                Data = tmp
43            Case Is > 2
44                Throw Err_Dimensions
45        End Select
          
46        ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))
          
47        If WriteToFile Then
48            If UCase$(Encoding) = "UTF-8" Then
49                Set Stream = CreateObject("ADODB.Stream")
50                Stream.Open
51                Stream.Type = 2 'Text
52                Stream.CharSet = "utf-8"
          
53                For i = LBound(Data) To UBound(Data)
54                    For j = LBound(Data, 2) To UBound(Data, 2)
55                        OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",", TrueString, FalseString)
56                    Next j
57                    OneLineJoined = VBA.Join(OneLine, Delimiter) & EOL
58                    Stream.WriteText OneLineJoined
59                Next i
60                Stream.SaveToFile FileName, 2 'adSaveCreateOverWrite

61                sCSVWrite = FileName
62            Else
63                Unicode = UCase$(Encoding) = "UTF-16"
64                If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
65                Set Stream = m_FSO.CreateTextFile(FileName, True, Unicode)
        
66                For i = LBound(Data) To UBound(Data)
67                    For j = LBound(Data, 2) To UBound(Data, 2)
68                        OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",", TrueString, FalseString)
69                    Next j
70                    OneLineJoined = VBA.Join(OneLine, Delimiter)
71                    WriteLineWrap Stream, OneLineJoined, EOLIsWindows, EOL, Unicode
72                Next i

73                Stream.Close: Set Stream = Nothing
74                sCSVWrite = FileName
75            End If
76        Else

77            ReDim Lines(LBound(Data) To UBound(Data) + 1) 'add one to ensure that result has a terminating EOL
        
78            For i = LBound(Data) To UBound(Data)
79                For j = LBound(Data, 2) To UBound(Data, 2)
80                    OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",", TrueString, FalseString)
81                Next j
82                Lines(i) = VBA.Join(OneLine, Delimiter)
83            Next i
84            sCSVWrite = VBA.Join(Lines, EOL)
85            If Len(sCSVWrite) > 32767 Then
86                If TypeName(Application.Caller) = "Range" Then
87                    Throw "Cannot return string of length " & Format$(CStr(Len(sCSVWrite)), "#,###") & _
                          " to a cell of an Excel worksheet"
88                End If
89            End If
90        End If
          
91        Exit Function
ErrHandler:
92        ErrRet = "#sCSVWrite: " & Err.Description & "!"
93        If Not Stream Is Nothing Then
94            Stream.Close
95            Set Stream = Nothing
96        End If
97        If m_ErrorStyle = es_ReturnString Then
98            sCSVWrite = ErrRet
99        Else
100           Throw ErrRet
101       End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(ByVal OS As String, ByVal ArgName As String) As String

          Const Err_Invalid As String = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associated end of line characters."

1         On Error GoTo ErrHandler
2         Select Case LCase$(OS)
              Case "windows", vbCrLf, "crlf"
3                 OStoEOL = vbCrLf
4             Case "unix", "linux", vbLf, "lf"
5                 OStoEOL = vbLf
6             Case "mac", vbCr, "cr"
7                 OStoEOL = vbCr
8             Case Else
9                 Throw ArgName & Err_Invalid
10        End Select

11        Exit Function
ErrHandler:
12        Throw "#OStoEOL: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of sCSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(ByVal x As Variant, ByVal QuoteAllStrings As Boolean, ByVal DateFormat As String, ByVal DateTimeFormat As String, _
          ByVal Delim As String, TrueString As String, FalseString As String) As String
          
          Const DQ As String = """"
          Const DQ2 As String = """"""

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 If InStr(x, DQ) > 0 Then
4                     Encode = DQ & Replace(x, DQ, DQ2) & DQ
5                 ElseIf QuoteAllStrings Then
6                     Encode = DQ & x & DQ
7                 ElseIf InStr(x, vbCr) > 0 Then
8                     Encode = DQ & x & DQ
9                 ElseIf InStr(x, vbLf) > 0 Then
10                    Encode = DQ & x & DQ
11                ElseIf InStr(x, Delim) > 0 Then
12                    Encode = DQ & x & DQ
13                Else
14                    Encode = x
15                End If
16            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbEmpty  'vbLongLong - not available on 16 bit.
17                Encode = CStr(x)
18            Case vbBoolean
19                Encode = IIf(x, TrueString, FalseString)
20            Case vbDate
21                If CLng(x) = CDbl(x) Then
22                    Encode = Format$(x, DateFormat)
23                Else
24                    Encode = Format$(x, DateTimeFormat)
25                End If
26            Case vbNull
27                Encode = "NULL"
28            Case vbError
29                Select Case CStr(x) 'Editing this case statement? Edit also its inverse, see method MakeSentinels
                      Case "Error 2000"
30                        Encode = "#NULL!"
31                    Case "Error 2007"
32                        Encode = "#DIV/0!"
33                    Case "Error 2015"
34                        Encode = "#VALUE!"
35                    Case "Error 2023"
36                        Encode = "#REF!"
37                    Case "Error 2029"
38                        Encode = "#NAME?"
39                    Case "Error 2036"
40                        Encode = "#NUM!"
41                    Case "Error 2042"
42                        Encode = "#N/A"
43                    Case "Error 2043"
44                        Encode = "#GETTING_DATA!"
45                    Case "Error 2045"
46                        Encode = "#SPILL!"
47                    Case "Error 2046"
48                        Encode = "#CONNECT!"
49                    Case "Error 2047"
50                        Encode = "#BLOCKED!"
51                    Case "Error 2048"
52                        Encode = "#UNKNOWN!"
53                    Case "Error 2049"
54                        Encode = "#FIELD!"
55                    Case "Error 2050"
56                        Encode = "#CALC!"
57                    Case Else
58                        Encode = CStr(x)        'should never hit this line...
59                End Select
60            Case Else
61                Throw "Cannot convert variant of type " & TypeName(x) & " to String"
62        End Select
63        Exit Function
ErrHandler:
64        Throw "#Encode: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write illegal characters to a stream opened with
'              TriStateFalse.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(t As TextStream, ByVal text As String, ByVal EOLIsWindows As Boolean, ByVal EOL As String, ByVal Unicode As Boolean)

          Dim ErrDesc As String
          Dim ErrNum As Long
          Dim i As Long

1         On Error GoTo ErrHandler
2         If EOLIsWindows Then
3             t.WriteLine text
4         Else
5             t.Write text
6             t.Write EOL
7         End If

8         Exit Sub

ErrHandler:
9         ErrNum = Err.Number
10        ErrDesc = Err.Description
11        If Not Unicode Then
12            If ErrNum = 5 Then
13                For i = 1 To Len(text)
14                    If Not CanWriteCharToAscii(Mid$(text, i, 1)) Then
15                        ErrDesc = "Data contains characters that cannot be written to an ascii file (first found is '" & _
                              Mid$(text, i, 1) & "' with unicode character code " & AscW(Mid$(text, i, 1)) & _
                              "). Try calling sCSVWrite with argument Encoding as ""UTF-8"" or ""UTF-16"""
16                        Exit For
17                    End If
18                Next i
19            End If
20        End If
21        Throw "#WriteLineWrap: " & ErrDesc & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CanWriteCharToAscii
' Purpose    : Not all characters for which AscW(c) < 255 can be written to an ascii file. If AscW(c) is in the following
'              list then they cannot:
'             128,130,131,132,133,134,135,136,137,138,139,140,142,145,146,147,148,149,150,151,152,153,154,155,156,158,159
' -----------------------------------------------------------------------------------------------------------------------
Private Function CanWriteCharToAscii(ByVal c As String) As Boolean
          Dim code As Long
1         code = AscW(c)
2         If code > 255 Or code < 0 Then
3             CanWriteCharToAscii = False
4         Else
5             CanWriteCharToAscii = Chr$(AscW(c)) = c
6         End If
End Function
