Attribute VB_Name = "modCSV_OLD"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Author    : Philip Swannell
' Date      : 26-Jul-2021, as from 22 Sep 2021 this function merely wraps sCSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' Delimiter : Delimiter character. Defaults to the first instance of comma, tab, semi-colon, colon or
'             vertical bar found in the file outside quoted regions. Enter FALSE to to see
'             the file's raw contents as would be displayed in a text editor.
' ShowNumbersAsNumbers: If TRUE, then numeric fields will be returned as numbers, otherwise as strings. This
'             argument is optional, defaulting to FALSE..
' ShowDatesAs: How date fields in the file are returned: 'String' (or FALSE) = return the string as is,
'             'Date' = convert to Date, 'Number` (or TRUE) = convert to number. This
'             argument is optional, defaulting to 'String'. See also argument DateFormat.
' ShowLogicalsAsLogicals: If TRUE, then fields in the file that are either TRUE or FALSE (case insensitive) will be
'             returned as Booleans, otherwise as strings. This argument is optional,
'             defaulting to False.
' LineEndings: "Windows" (or ascii 13 + 10, or FALSE), "Unix" (or ascii 10, or TRUE) or "Mac" (or ascii
'             13)  to specify line-endings, or "Mixed" for inconsistent line endings. Omit
'             to infer from the first line ending found outside quoted regions.
' DateFormat: The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read
'             from Windows regional settings. Repeated D's (or M's or Y's) are equivalent
'             to single instances, so that d-m-y and DD-MMM-YYYY are equivalent.
' DecimalSeparator: The character that represents a decimal point. If omitted, then the value from Windows
'             regional settings is used.
' ThousandsSeparator: The character used as a thousands separator when representing numbers. If omitted, then
'             the value from Windows regional settings is used.
' ShowErrorsAsErrors: If TRUE, then fields in the file that appear as Excel-style errors such as #NAME? or #REF!
'             will be converted to error values, otherwise they will be returned as
'             strings. This argument is optional and defaults to False.
' StartRow  : The "row" in the file at which reading starts. Optional and defaults to 1 to read from the
'             first row.
' StartCol  : The "column" in the file at which reading starts. Optional and defaults to 1 to read from
'             the first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from StartRow to
'             the end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from
'             StartColumn are read.
' Unicode   : Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to guess (via function
'             sFileIsUnicode).
' ShowMissingsAs: Value to represent empty fields (successive delimiters) in the file. May be a string or an
'             Empty value. Optional and defaults to the zero-length string.
' RemoveQuotes: If TRUE (the default) then correctly (RFC4180) quoted fields are unquoted. I.e. fields
'             which both start and end with double quotes have those two characters removed
'             and adjacent pairs of double quotes are replaced by single double quotes.
'
' Notes     : See also sFileSaveCSV for which this function is the inverse.
'
'             The function handles all csv files that conform to the standards described in
'             RFC4180  https://www.rfc-editor.org/rfc/rfc4180.txt including files with
'             quoted fields.
'
'             In addition the function handles files which break some of those standards:
'             * Not all lines of the file need have the same number of fields. The function
'             "pads" with the value given by ShowMissingsAs.
'             * Fields which start with a double quote but do not end with a double quote
'             are handled by being returned unchanged. Necessarily such fields have an even
'             number of double quotes, or otherwise the field will be treated as the last
'             field in the file.
'             * The standard states that csv files should have Windows-style line endings,
'             but the function supports Windows, Unix and (old) Mac line endings.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileShow(FileName As String, Optional ByVal Delimiter As Variant, Optional ShowNumbersAsNumbers As Boolean, _
          Optional ShowDatesAs As Variant = "String", Optional ShowLogicalsAsLogicals As Boolean, _
          Optional ByVal LineEndings As Variant, Optional ByVal DateFormat As String, _
          Optional DecimalSeparator As String = vbNullString, Optional ThousandsSeparator As String = vbNullString, _
          Optional ShowErrorsAsErrors As Boolean, Optional ByVal StartRow As Long = 1, Optional ByVal StartCol As Long = 1, _
          Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, Optional ByVal Unicode As Variant, _
          Optional ByVal ShowMissingsAs As Variant = "", Optional ByVal RemoveQuotes As Boolean = True)
Attribute sFileShow.VB_Description = "DEPRECATED FUNCTION.\nUse sCSVRead instead.\nReturns the contents of a comma-separated file on disk as an array."
Attribute sFileShow.VB_ProcData.VB_Invoke_Func = " \n26"
                
1         On Error GoTo ErrHandler
          
          Dim ConvertTypes As String
          Dim Encoding As Variant
          Const IgnoreRepeated = False
          Const Comment = ""
          Const IgnoreEmptyLines = False
          Const HeaderRowNum = 0
          
          'Convert inputs to be suitable for sCSVRead
2         If IsMissing(Unicode) Then
3             Encoding = CreateMissing()
4         ElseIf VarType(Unicode) = vbBoolean Then
5             Encoding = IIf(Unicode, "UTF-16", "ANSI")
6         Else
7             Throw "Unicode must be TRUE, FALSE or omitted"
8         End If
          'for backward compatibility
9         If LCase(DateFormat) = "false" Then
10            DateFormat = ""
11        End If
          
          'For function sFileShow Delimiter = "" meant "file is not delimited", for sCSVRead we must pass Delimter = FALSE
12        If VarType(Delimiter) = vbString Then
13            If Delimiter = "" Then
14                Delimiter = False
15            End If
16        End If
          
17        ConvertTypes = IIf(ShowNumbersAsNumbers, "N", "") & IIf(LCase(ShowDatesAs) = "string", "", "D") & _
              IIf(ShowLogicalsAsLogicals, "B", "") & IIf(ShowErrorsAsErrors, "E", "")
              
18        If Not RemoveQuotes Then Throw "RemoveQuotes = False is no longer supported"
          
19        sFileShow = sCSVRead(FileName, ConvertTypes, Delimiter, IgnoreRepeated, DateFormat, Comment, _
              IgnoreEmptyLines, HeaderRowNum, StartRow, StartCol, NumRows, NumCols, , , , ShowMissingsAs, Encoding, DecimalSeparator)
          
20        Exit Function
ErrHandler:
21        sFileShow = "#sFileShow (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    :
' Parameters :
'  DateFormat   : String such as D/M/Y or Y-M-D.
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/"
' -----------------------------------------------------------------------------------------------------------------------
Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String)

          Const Err_DateFormat = "DateFormat should be 'M-D-Y', 'D-M-Y' or 'Y-M-D'. A character other " + _
              "than '-' is allowed as the separator. If there is no separator then the allowed formats are 'MMDDYYYY', 'DDMMYYYY' and 'YYYYMMDD'. Omit to use the Windows default, which on this PC is "

1         On Error GoTo ErrHandler
          
2         If Len(DateFormat) = 8 Then
3             Select Case LCase(DateFormat)
                  Case "mmddyyyy"
4                     DateOrder = 0
5                     DateSeparator = ""
6                     Exit Sub
7                 Case "ddmmyyyy"
8                     DateOrder = 1
9                     DateSeparator = ""
10                    Exit Sub
11                Case "yyyymmdd"
12                    DateOrder = 2
13                    DateSeparator = ""
14                    Exit Sub
15            End Select
16        End If
          
          'Replace repeated D's with a single D, etc since sParseDateCore only needs _
           to know the order in which the three parts of the date appear.

17        If Len(DateFormat) > 5 Then
18            DateFormat = UCase(DateFormat)
19            ReplaceRepeats DateFormat, "D"
20            ReplaceRepeats DateFormat, "M"
21            ReplaceRepeats DateFormat, "Y"
22        End If

23        If Len(DateFormat) = 0 Then
24            DateOrder = Application.International(xlDateOrder)
25            DateSeparator = Application.International(xlDateSeparator)
26        ElseIf Len(DateFormat) <> 5 Then
27            Throw Err_DateFormat + WindowsDefaultDateFormat
28        ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
29            Throw Err_DateFormat + WindowsDefaultDateFormat
30        Else
31            DateSeparator = Mid$(DateFormat, 2, 1)
32            Select Case UCase$(Left$(DateFormat, 1) + Mid$(DateFormat, 3, 1) + Right$(DateFormat, 1))
                  Case "MDY"
33                    DateOrder = 0
34                Case "DMY"
35                    DateOrder = 1
36                Case "YMD"
37                    DateOrder = 2
38                Case Else
39                    Throw Err_DateFormat + WindowsDefaultDateFormat
40            End Select
41        End If

42        Exit Sub
ErrHandler:
43        Throw "#ParseDateFormat (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
          Dim ChCh As String
1         ChCh = TheChar & TheChar
2         While InStr(TheString, ChCh) > 0
3             TheString = Replace(TheString, ChCh, TheChar)
4         Wend
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WindowsDefaultDateFormat
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Returns a description of the system date format, used only for error message generation.
' -----------------------------------------------------------------------------------------------------------------------
Private Function WindowsDefaultDateFormat() As String
          Dim DS As String
1         On Error GoTo ErrHandler
2         DS = Application.International(xlDateSeparator)
3         Select Case Application.International(xlDateOrder)
              Case 0
4                 WindowsDefaultDateFormat = "M" + DS + "D" + DS + "Y"
5             Case 1
6                 WindowsDefaultDateFormat = "D" + DS + "M" + DS + "Y"
7             Case 2
8                 WindowsDefaultDateFormat = "Y" + DS + "M" + DS + "D"
9         End Select

10        Exit Function
ErrHandler:
11        WindowsDefaultDateFormat = "Cannot determine!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileSaveCSV
' Author    : Philip Swannell
' Date      : 18-Dec-2015, as from 22 Sep 2021 this function merely wraps sCSVWrite
' Purpose   : Creates a csv file on disk containing the data in the array Data. Any existing file of the
'             same name is overwritten. If successful, the function returns the name of the
'             file written, otherwise an error string.
' Arguments
' FileName  : The full name of the file, including the path.
' Data      : An array of arbitrary data. Elements may be strings, numbers, Booleans, empty, errors or
'             null values.
' QuoteAllStrings: If TRUE (the default) then ALL strings are quoted before being written to file. Otherwise
'             (FALSE) only strings containing characters comma, line feed, carriage return
'             or double quote are quoted. Double quotes are always escaped by a second
'             double quote.
' DateFormat: A format string, such as 'yyyy-mm-dd' that determine how dates (e.g. cells formatted as
'             dates) appear in the file.
' DateTimeFormat: A format string, such as 'yyyy-mm-dd hh:mm:ss' that determine how elements of dates with
'             time appear in the file.
' Unicode   : If TRUE then the file written is encoded as unicode. Defaults to FALSE for an ascii file.
'
' Notes     : See also sFileShow which is the inverse of this function.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSaveCSV(FileName As String, ByVal Data As Variant, Optional QuoteAllStrings As Boolean = True, _
          Optional DateFormat As String = "yyyy-mm-dd", Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
          Optional Unicode As Boolean, Optional ByVal EOL = vbCrLf)
Attribute sFileSaveCSV.VB_Description = "DEPRECATED FUNCTION.\nUse sCSVWrite instead.\n\nCreates a csv file on disk containing the data in the array Data. Any existing file of the same name is overwritten. If successful, the function returns the name of the file written, otherwise an error string. "
Attribute sFileSaveCSV.VB_ProcData.VB_Invoke_Func = " \n26"
                
1         sFileSaveCSV = sCSVWrite(Data, FileName, QuoteAllStrings, DateFormat, DateTimeFormat, ",", IIf(Unicode, "UTF-16", "ANSI"), EOL)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsUnicodeFile
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : Tests if a file is unicode by reading the byte-order-mark. Return is True, False or an error string, so
'              calls should usually be wrapped in ThrowIfError. Adapted from
'              https://stackoverflow.com/questions/36188224/vba-test-encoding-of-a-text-file
' -----------------------------------------------------------------------------------------------------------------------
Public Function IsUnicodeFile(FilePath As String)
          Static FSO As Scripting.FileSystemObject
          Dim t As Scripting.TextStream

          Dim intAsc1Chr As Long
          Dim intAsc2Chr As Long

1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
3         If (FSO.FileExists(FilePath) = False) Then
4             IsUnicodeFile = "#File not found!"
5             Exit Function
6         End If

          ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
7         Set t = FSO.OpenTextFile(FilePath, 1, False, 0)
8         If t.atEndOfStream Then
9             t.Close: Set t = Nothing
10            IsUnicodeFile = False
11            Exit Function
12        End If
13        intAsc1Chr = Asc(t.Read(1))
14        If t.atEndOfStream Then
15            t.Close: Set t = Nothing
16            IsUnicodeFile = False
17            Exit Function
18        End If
19        intAsc2Chr = Asc(t.Read(1))
20        t.Close
21        If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then 'UTF-16 LE BOM
22            IsUnicodeFile = True
23        Else
24            IsUnicodeFile = False
25        End If

26        Exit Function
ErrHandler:
27        IsUnicodeFile = "#IsUnicodeFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
