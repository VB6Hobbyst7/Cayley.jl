Attribute VB_Name = "modFileFnsB"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFilesMerge
' Author    : Philip Swannell
' Date      : 19-Mar-2019
' Purpose   : Element-wise "merging" of many comma-delimited text files. The (i,j)th element of the
'             output file is the result of the operator function (such as Median) on the
'             set of (i,j)th elements of the input files.
' Arguments
' InputFiles: An array of file names (with path) for the input files.
' OutputFile: The file name (with path) for the output file.
' Delimiter : The delimiter character used for both the input files and the output file.
' NumTopHeaders: The number of header rows at the top of each input file. If positive then header rows must
'             be identical in all input files and appear in the output file.
' NumLeftHeaders: The number of header columns at the top of each input file. If positive then header
'             columns must be identical in all input files and appear in the output file.
' Operator  : Enter text to specify the merging function. Allowed values "Median", "InterQuartileRange"
'             and "CountNumeric". InterQuartileRange is calculated using the Excel function
'             QUARTILE.INC.
' AllowMismatchedTopHeaders: If the top headers of the input files are not identical set this argument to TRUE and the
'             function will generate a file with top headers equal to the union of the top
'             headers in the input files.
' TopHeaderRightDelimiter: If supplied, then when top headers ar read from each file, only characters before the
'             first occurrence of this character are read, all later characters are
'             ignored.
' -----------------------------------------------------------------------------------------------------------------------
Function sFilesMerge(InputFiles As Variant, ByVal OutputFiles As Variant, Delimiter As String, NumTopHeaders As Long, _
        NumLeftHeaders As Long, Optional ByVal Operators As Variant = "Median", Optional AllowMismatchedTopHeaders As Boolean, _
        Optional TopHeaderRightDelimiter As String, Optional SuppressZeros As Boolean)
Attribute sFilesMerge.VB_Description = "Element-wise ""merging"" of many comma-delimited text files. The (i,j)th element of the output file is the result of the operator function (such as Median) on the set of (i,j)th elements of the input files."
Attribute sFilesMerge.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim DataToWriteM, DataToWriteIQR, DataToWriteCN, DataToWriteTemplate
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NFiles As Long
          Dim NR As Long
          Dim ThreeDArray() As Variant
          Dim DoMedian As Boolean
          Dim DoIQR As Boolean
          Dim DoCN As Boolean
          
          Dim OpCodes() As Long    '1 = Median, 2 = InterQuartileRange, 3 = CountNumeric

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti InputFiles, OutputFiles, Operators

          'check OutputFiles and Operators conformant
3         If sNCols(OutputFiles) <> 1 Or sNCols(Operators) <> 1 Or (sNRows(OutputFiles) <> sNRows(Operators)) Then
4             Throw "OutputFiles and Operators must be strings or single column arrays with the same number of rows"
5         End If

          'get an error early if don't have access to write output files
          Dim Res
6         For i = 1 To sNRows(OutputFiles)
7             If sFileExists(OutputFiles(i, 1)) Then
8                 Res = sFileDelete(CStr(OutputFiles(i, 1)))
9                 If sIsErrorString(Res) Then
10                    Throw "Output file '" + OutputFiles(i, 1) + "' already exists and can't be deleted - " + Res
11                End If
12            Else
13                Res = sFileSave(CStr(OutputFiles(i, 1)), "Test data", ",")
14                If sIsErrorString(Res) Then
15                    Throw "Test data cannot be written to '" + OutputFiles(i, 1) + "' - " + Res
16                End If
17                ThrowIfError sFileDelete(CStr(OutputFiles(i, 1)))
18            End If
19        Next i

20        InputFiles = sReshape(InputFiles, sNRows(InputFiles) * sNCols(InputFiles), 1)
21        NFiles = sNRows(InputFiles)
22        If NumLeftHeaders < 0 Then Throw "NumLeftHeaders must be greater than or equal to zero"
23        If NumTopHeaders < 0 Then Throw "NumTopHeaders must be greater than or equal to zero"

24        ReDim OpCodes(1 To sNRows(Operators))

25        For i = 1 To sNRows(Operators)
26            Select Case Replace(Replace(LCase$(Operators(i, 1)), " ", vbNullString), "-", vbNullString)
                  Case "median", "m"
27                    OpCodes(i) = 1
28                    DoMedian = True
29                Case "interquartilerange", "iqr"
30                    OpCodes(i) = 2
31                    DoIQR = True
32                Case "countnumeric", "cn", "countnumbers", "countnumber"
33                    OpCodes(i) = 3
34                    DoCN = True
35                Case Else
36                    Throw "Operator not recognised. Allowed values are 'Median', 'Inter Quartile Range' and 'Count Numeric' or a column array of such."
37            End Select
38        Next

39        MultipleFilesTo3DArray ThreeDArray, InputFiles, Delimiter, NumTopHeaders, NumLeftHeaders, AllowMismatchedTopHeaders, TopHeaderRightDelimiter

          Dim BlankArray() As Variant
40        ReDim BlankArray(1 To NFiles)
41        NR = UBound(ThreeDArray, 1)
42        NC = UBound(ThreeDArray, 2)

43        DataToWriteTemplate = sReshape(vbNullString, NR, NC)
44        For i = 1 To NumTopHeaders
45            For j = 1 To NC
46                DataToWriteTemplate(i, j) = ThreeDArray(i, j, 1)
47            Next j
48        Next i
49        For j = 1 To NumLeftHeaders
50            For i = 1 To NR
51                DataToWriteTemplate(i, j) = ThreeDArray(i, j, 1)
52            Next i
53        Next j

54        If DoMedian Then
55            DataToWriteM = DataToWriteTemplate
56        End If
57        If DoIQR Then
58            DataToWriteIQR = DataToWriteTemplate
59        End If
60        If DoCN Then
61            DataToWriteCN = DataToWriteTemplate
62        End If

          Dim NumNums As Long
63        For i = NumTopHeaders + 1 To NR
64            For j = NumLeftHeaders + 1 To NC
65                ReDim BlankArray(1 To NFiles)
                  'Thanks to the args passed to sFileShow we can be sure that all elements of ThreeDArray are numbers or strings
66                NumNums = 0
67                For k = 1 To NFiles
68                    If VarType(ThreeDArray(i, j, k)) = vbString Then
69                        BlankArray(k) = ""
70                    Else
71                        BlankArray(k) = ThreeDArray(i, j, k)
72                        NumNums = NumNums + 1
73                        If SuppressZeros Then
74                            If BlankArray(k) = 0 Then
75                                BlankArray(k) = ""
76                                NumNums = NumNums - 1
77                            End If
78                        End If
79                    End If
80                Next k
81                If DoMedian Then
82                    If NumNums > 0 Then
83                        DataToWriteM(i, j) = Application.WorksheetFunction.Median(BlankArray)
84                    Else
85                        DataToWriteM(i, j) = "NA"
86                    End If
87                End If

88                If DoIQR Then
89                    If NumNums > 0 Then
90                        DataToWriteIQR(i, j) = SafeIRQ(BlankArray)
91                    Else
92                        DataToWriteIQR(i, j) = "NA"
93                    End If
94                End If
95                If DoCN Then
96                    DataToWriteCN(i, j) = NumNums
97                End If

98            Next j
99        Next i

100       For i = 1 To sNRows(OutputFiles)
101           Select Case OpCodes(i)
                  Case 1
102                   ThrowIfError sFileSave(CStr(OutputFiles(i, 1)), DataToWriteM, Delimiter)
103               Case 2
104                   ThrowIfError sFileSave(CStr(OutputFiles(i, 1)), DataToWriteIQR, Delimiter)
105               Case 3
106                   ThrowIfError sFileSave(CStr(OutputFiles(i, 1)), DataToWriteCN, Delimiter)
107           End Select
108       Next i

109       sFilesMerge = OutputFiles

110       Exit Function
ErrHandler:
111       sFilesMerge = "#sFilesMerge (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MultipleFilesTo3DArray
' Author     : Philip Swannell
' Date       : 22-Apr-2020
' Purpose    : Much of the functionality of sFilesMerge is delegated to this method which is also called from code of
'              workbook "ISDA SIMM YYYY Analyse Data Series.xlsm"
' Parameters :
'  ThreeDArray              : The return value, passed by reference. "Layer" k of the return is the contents of file(i)
'                             but "appropriately morphed" to align columns in the event that the top headers are in a
'                             different order or the set-difference of the top headers is non empty.
'  InputFiles               : A column array of input files, (typically csv). All should have identical left headers
'                             and in the simple case all have identical top headers.
'  Delimiter                :
'  NumTopHeaders            :
'  NumLeftHeaders           :
'  AllowMismatchedTopHeaders: Should non-matching of top headers be an error or should we cope?
'  TopHeaderRightDelimiter  : For ISDASIMM work the top headers in each file have a postfix unique to that file which needs
'                             to be removed prior to matching the headers between files. So the last occurence of this character
'                             and characters to the right are stripped when reading the input files
' -----------------------------------------------------------------------------------------------------------------------
Sub MultipleFilesTo3DArray(ByRef ThreeDArray As Variant, ByRef InputFiles As Variant, Delimiter As String, NumTopHeaders As Long, _
          NumLeftHeaders As Long, Optional AllowMismatchedTopHeaders As Boolean, _
          Optional TopHeaderRightDelimiter As String, Optional LogProgress As Boolean)
          
          Dim FirstLeftHeaders
          Dim FirstTopHeaders
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NFiles As Long
          Dim NR As Long
          Dim ThisFileContents As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR InputFiles

3         InputFiles = sReshape(InputFiles, sNRows(InputFiles) * sNCols(InputFiles), 1)
4         NFiles = sNRows(InputFiles)
5         If NumLeftHeaders < 0 Then Throw "NumLeftHeaders must be greater than or equal to zero"
6         If NumTopHeaders < 0 Then Throw "NumTopHeaders must be greater than or equal to zero"

7         If AllowMismatchedTopHeaders Then
8             If NumTopHeaders < 1 Then Throw "NumTopHeaders must be positive when AllowMismatchedTopHeaders is TRUE"
              Dim MatchIDs As Variant
              Dim OutputTopHeaders As Variant
              Dim STK As clsStacker
              Dim ThisTopHeaders As Variant
              
9             OutputTopHeaders = ThrowIfError(sFileShow(CStr(InputFiles(1, 1)), Delimiter, True, , , , , , , , 1, 1, NumTopHeaders))
10            If TopHeaderRightDelimiter <> vbNullString Then MorphTopHeaders OutputTopHeaders, NumTopHeaders, TopHeaderRightDelimiter
11            Set STK = CreateStacker()
12            STK.Stack2D sArrayTranspose(OutputTopHeaders)
13            For k = 2 To NFiles
14                ThisTopHeaders = ThrowIfError(sFileShow(CStr(InputFiles(k, 1)), Delimiter, True, , , , , , , , 1, 1, NumTopHeaders))
15                If TopHeaderRightDelimiter <> vbNullString Then MorphTopHeaders ThisTopHeaders, NumTopHeaders, TopHeaderRightDelimiter
16                If NumLeftHeaders > 1 Then
17                    If Not sArraysIdentical(sSubArray(OutputTopHeaders, 1, 1, , NumLeftHeaders), sSubArray(ThisTopHeaders, 1, 1, , NumLeftHeaders)) Then
18                        Throw "Even if AllowMismatchedTopHeaders is TRUE, the top-left part of the files measuring NumLeftHeaders by NumTopHeaders must be identical but that of file 1 does not match that of file " + CStr(k)
19                    End If
20                End If
21                If sNCols(ThisTopHeaders) > NumLeftHeaders Then
22                    STK.Stack2D sArrayTranspose(sSubArray(ThisTopHeaders, 1, NumLeftHeaders + 1))
23                End If
24            Next k
25            OutputTopHeaders = sArrayTranspose(sRemoveDuplicateRows(STK.Report, False))
26        End If

27        For k = 1 To NFiles
28            If LogProgress Then MessageLogWrite "MultipleFilesTo3DArray: Reading file '" + CStr(InputFiles(k, 1)) + "'"
29            ThisFileContents = sFileShow(CStr(InputFiles(k, 1)), Delimiter, True)
30            If sIsErrorString(ThisFileContents) Then
31                Throw "Error reading file " + CStr(k) + " - " + ThisFileContents
32            End If
              
33            If TopHeaderRightDelimiter <> vbNullString Then MorphTopHeaders ThisFileContents, NumTopHeaders, TopHeaderRightDelimiter

34            If AllowMismatchedTopHeaders Then
                  'Morph ThisFileContents so that the columns align with the output columns
                  Dim Result
35                ThisTopHeaders = sSubArray(ThisFileContents, 1, 1, NumTopHeaders)
36                If sNCols(ThisTopHeaders) <> sNRows(sRemoveDuplicateRows(sArrayTranspose(ThisTopHeaders))) Then
37                    Throw "If AllowMismatchedTopHeaders is TRUE then TopHeaders in each file must not contain repeats, but there are repeats in the TopHeaders of '" + CStr(InputFiles(k, 1)) + "'"
38                End If
39                If Not sArraysIdentical(ThisTopHeaders, OutputTopHeaders) Then
40                    If NumTopHeaders = 1 Then
41                        MatchIDs = sMatch(sArrayTranspose(ThisTopHeaders), sArrayTranspose(OutputTopHeaders))
42                    Else
43                        MatchIDs = sMultiMatch(sArrayTranspose(ThisTopHeaders), sArrayTranspose(OutputTopHeaders), False)
44                    End If
                      'sMatch is what in Julia would be called type-unstable. So here we force to 2d array
45                    If sNRows(MatchIDs) = 1 Then
46                        Force2DArray MatchIDs
47                    End If

48                    Result = sArrayStack(OutputTopHeaders, sReshape(" ", sNRows(ThisFileContents) - NumTopHeaders, sNCols(OutputTopHeaders)))
49                    For j = 1 To sNCols(ThisTopHeaders)
50                        If IsNumber(MatchIDs(j, 1)) Then
51                            For i = NumTopHeaders + 1 To sNRows(ThisFileContents)
52                                Result(i, MatchIDs(j, 1)) = ThisFileContents(i, j)
53                            Next i
54                        End If
55                    Next j
56                    ThisFileContents = Result
57                End If
58            End If

59            If k = 1 Then
60                NR = sNRows(ThisFileContents)
61                NC = sNCols(ThisFileContents)
62                If NumTopHeaders > 0 Then
63                    FirstTopHeaders = sSubArray(ThisFileContents, 1, 1, NumTopHeaders)
64                End If
65                If NumLeftHeaders > 0 Then
66                    FirstLeftHeaders = sSubArray(ThisFileContents, 1, 1, , NumLeftHeaders)
67                End If
68                ReDim ThreeDArray(1 To NR, 1 To NC, 1 To NFiles)
69            Else
70                If sNRows(ThisFileContents) <> NR Then
71                    Throw "Files must all have the same number of rows, but file 1 has " + CStr(NR) + " row(s) and file " + CStr(k) + " has + " + CStr(sNRows(ThisFileContents)) + " row(s)"
72                End If
73                If sNCols(ThisFileContents) <> NC Then
74                    Throw "Files must all have the same number of columns, but file 1 has " + CStr(NC) + " column(s) and file " + CStr(k) + " has + " + CStr(sNCols(ThisFileContents)) + " column(s)"
75                End If
76                If NumTopHeaders > 0 Then
77                    If Not sArraysIdentical(FirstTopHeaders, sSubArray(ThisFileContents, 1, 1, NumTopHeaders)) Then
78                        Throw "The top header row(s) in files 1 and " + CStr(k) + " are not identical. Try setting argument AllowMismatchedTopHeaders to TRUE"
79                    End If
80                End If
81                If NumLeftHeaders > 0 Then
82                    If Not ArraysIdenticalSpecial(sDrop(FirstLeftHeaders, NumTopHeaders), sSubArray(ThisFileContents, NumTopHeaders + 1, 1, , NumLeftHeaders)) Then
83                        Throw "The left header column(s) in files 1 and " + CStr(k) + " are not identical"
84                    End If
85                End If
86            End If
87            For i = 1 To NR
88                For j = 1 To NC
89                    ThreeDArray(i, j, k) = ThisFileContents(i, j)
90                Next j
91            Next i
92        Next k
93        If LogProgress Then MessageLogWrite "MultipleFilesTo3DArray: finished reading " + CStr(NFiles) + " files."

94        Exit Sub
ErrHandler:
95        Throw "#MultipleFilesTo3DArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'Tests if arrays identical allowing that they might both be arrays of dates represented as strings using different guessable formats
Private Function ArraysIdenticalSpecial(Array1, Array2) As Boolean
1         On Error GoTo ErrHandler
2         If sArraysIdentical(Array1, Array2) Then
3             ArraysIdenticalSpecial = True
4             Exit Function
5         ElseIf sNRows(Array1) <> sNRows(Array2) Then
6             ArraysIdenticalSpecial = False
7             Exit Function
8         ElseIf sNCols(Array1) <> sNCols(Array2) Then
9             ArraysIdenticalSpecial = False
10            Exit Function
11        Else
              Dim Format1 As String, Format2 As String
12            Format1 = ISDASIMMGuessDateFormat(Array1)
13            If Not sIsErrorString(Format1) Then
14                Format2 = ISDASIMMGuessDateFormat(Array2)
15                If Not sIsErrorString(Format2) Then
16                    ArraysIdenticalSpecial = sArraysIdentical(sParseDate(Array1, Format1), sParseDate(Array2, Format2))
17                    Exit Function
18                End If
19            End If
20        End If
21        ArraysIdenticalSpecial = False
22        Exit Function
ErrHandler:
23        Throw "#ArraysIdenticalSpecial (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub MorphTopHeaders(ByRef FileContents, NumTopHeaders As Long, RightDelimiter As String)
          Dim i As Long
          Dim j As Long
          Dim matchpoint As Long
1         On Error GoTo ErrHandler
2         For i = 1 To NumTopHeaders
3             For j = 1 To sNCols(FileContents)
4                 If VarType(FileContents(i, j)) = vbString Then
5                     matchpoint = InStrRev(FileContents(i, j), RightDelimiter)
6                     If matchpoint > 0 Then
7                         FileContents(i, j) = Left$(FileContents(i, j), matchpoint - 1)
8                     End If
9                 End If
10            Next j
11        Next i
12        Exit Sub
ErrHandler:
13        Throw "#MorphTopHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function SafeIRQ(Data)
1         On Error GoTo ErrHandler
2         SafeIRQ = Application.WorksheetFunction.Quartile_Inc(Data, 3) - Application.WorksheetFunction.Quartile_Inc(Data, 1)
3         Exit Function
ErrHandler:
4         SafeIRQ = "NA"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileSplit
' Author    : Philip Swannell
' Date      : 27-Jul-2017
' Purpose   : Reads a large text file and splits its contents into a number of smaller files. The first
'             NumLinesPerFile lines are written to the first file, the next NumLinesPerFile
'             lines are written to the second, etc.
' Arguments
' InputFile : The name (with path) of the text file to read.
' NumLinesPerFile: The number of lines to be written to each of the files created. The last file created is
'             likely to have fewer lines.
' RepeatHeaders: If TRUE, then the first line of InputFile is appears as the first line of each file
'             created. If FALSE, then the first line of InputFile appears only in the first
'             file created.
' FileTemplate: Specifies the names of the files created. Must include left and right-angle braces <>
'             which are replaced by an integer count. Example c:\temp\OutputFile<>.csv
'
' Notes     : The function returns a column array of file names created.
'
'             The function has been tested on an input file containing 48 million lines and
'             created 48 one-million-line files in under two minutes.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSplit(InputFile As String, NumLinesPerFile As Long, RepeatHeaders As Boolean, FileTemplate As String)
Attribute sFileSplit.VB_Description = "Reads a large text file and splits its contents into a number of smaller files. The first NumLinesPerFile lines are written to the first file, the next NumLinesPerFile lines are written to the second, etc."
Attribute sFileSplit.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim ArrayIndex As Long
          Dim ChunkSize As Long
          Dim FileContents() As String
          Dim FileCounter As Long
          Dim fIn As Scripting.TextStream
          Dim fOut As Scripting.TextStream
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim NeedToWrite As Boolean
          Dim outputFolder As String
          Dim ReadCounter As Long
          Dim Result() As String
          Dim ShortFileTemplate As String
          Dim ThisFileName As String
          Dim ThisLine As String
          Dim TopLine As String

1         On Error GoTo ErrHandler
2         If Not sFileExists(InputFile) Then Throw "Cannot find file '" + InputFile + "'!"
3         outputFolder = sSplitPath(FileTemplate, False)
4         If Not sFolderIsWritable(outputFolder) Then Throw "Folder '" + outputFolder + "' does not exist or is not writeable"

5         ShortFileTemplate = sSplitPath(FileTemplate, True)
6         If InStr(ShortFileTemplate, "<>") = 0 Then Throw "FileTemplate must include characters ""<>"" to indicate position of file number"
7         If NumLinesPerFile < 2 Then Throw " NumLinesPerFile must be at least 2"

8         Set FSO = New FileSystemObject
9         Set fIn = FSO.OpenTextFile(InputFile, ForReading)

10        ReDim FileContents(1 To NumLinesPerFile)

11        If RepeatHeaders Then
12            ChunkSize = NumLinesPerFile - 1
13            TopLine = fIn.ReadLine
14            FileContents(1) = TopLine
15            ArrayIndex = 2
16        Else
17            ChunkSize = NumLinesPerFile
18            ArrayIndex = 1
19        End If

20        FileCounter = 0
21        Do While Not fIn.atEndOfStream
22            ThisLine = fIn.ReadLine
23            ReadCounter = ReadCounter + 1
24            FileContents(ArrayIndex) = ThisLine
25            NeedToWrite = True
26            ArrayIndex = ArrayIndex + 1
27            If ReadCounter Mod ChunkSize = 0 Then
28                FileCounter = FileCounter + 1
29                ThisFileName = outputFolder + "\" + Replace(ShortFileTemplate, "<>", CStr(FileCounter))
30                Set fOut = FSO.OpenTextFile(ThisFileName, ForWriting, True)
31                fOut.Write VBA.Join(FileContents, vbCrLf) + vbCrLf
32                NeedToWrite = False
33                fOut.Close: Set fOut = Nothing
34                ReDim FileContents(1 To NumLinesPerFile + 1)
35                ArrayIndex = 1
36                If FileCounter >= 1 And RepeatHeaders Then
37                    FileContents(1) = TopLine
38                    ArrayIndex = 2
39                End If
40            End If
41        Loop

42        If NeedToWrite Then
43            FileCounter = FileCounter + 1
44            ThisFileName = outputFolder + "\" + Replace(ShortFileTemplate, "<>", CStr(FileCounter))
45            Set fOut = FSO.OpenTextFile(ThisFileName, ForWriting, True)
46            ReDim Preserve FileContents(1 To ArrayIndex - 1)
47            fOut.Write VBA.Join(FileContents, vbCrLf) + vbCrLf
48            NeedToWrite = False
49            fOut.Close: Set fOut = Nothing
50        End If

51        fIn.Close: Set fIn = Nothing: Set FSO = Nothing

52        ReDim Result(1 To FileCounter, 1 To 1)
53        For i = 1 To FileCounter
54            Result(i, 1) = outputFolder + "\" + Replace(ShortFileTemplate, "<>", CStr(i))
55        Next i

56        sFileSplit = Result

57        Exit Function
ErrHandler:
58        sFileSplit = "#sFileSplit (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFileSplitColumns
' Author     : Philip Swannell
' Date       : 12-Apr-2019
' Purpose    : Reads a comma-separated text file and allocates the columns of that file to a number of output files. Where:
'            * The first column of the input file appears as the first column of each output file
'            * Subsequent columns of the input file are allocated to output files according to the contents of a "bucketing file"
'              (or array) that has two columns mapping (col1) the header text in the input file to a "bucket number" (or string)
'              the bucket number is shorthand for the name of the output file - see argument OutputFileTemplate.
'            * Each output file has a header row as its first row.
'            * If FirstColToAll is TRUE then the first column of the input file appears as the first column of all output files.
' Parameters :
'  InputFileName       : Full name of the input file, must be comma-delimited, and file parsing is quite naive, via VBA.Split
'  HeaderRowNumber     : The line number of the headers in the input file.
'  BucketingFileOrArray: The bucketing file or array, as described above. Not all headers in the input file need appear in the
'                        bucketing file, where they don't appear the corresponding column of the input file appears in no output file.
'  BucketingFileIsUnix : Boolean, passed to call to sFileShow
'  OutputFileTemplate  : Text that defines the name of the output file, each output file being obtained by replacing the
'                        characters "??" in the OutputFileTeplate with the relevant bucket number
'  BucketNumberFormat  : Where bucket numbers (right col of BucketingFile) are numbers then they are formatted with this format string
'                        before replacing the question marks in the OutputFileTemplate to yield the names of the output file.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSplitColumns(InputFileName As String, HeaderRowNumber As Long, BucketingFileOrArray As Variant, OutputFileTemplate As String, _
        FirstColToAll As Boolean, Optional BucketingFileIsUnix As Boolean, Optional BucketNumberFormat As String = "00")
Attribute sFileSplitColumns.VB_Description = "Reads a text file and splits its contents by column into output files. Columns are allocated according to a bucketing file (or array) and (if FirstColToAll is TRUE) the first column of the input file appears as the first column of each output file."
Attribute sFileSplitColumns.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim BucketsWrittenTo
          Dim ChooseVector
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim InputFileHeadersT
          Dim LookUpErrorText As String
          Dim NumTargetFiles As Long
          Dim t As Scripting.TextStream
          Dim TargetFileIndexes
          Dim TargetFileMonikers
          Dim TargetFileNames() As String
          Dim TargetTextStreams() As Scripting.TextStream
          Dim ThisBucket
          Dim ThisFileName As String

1         On Error GoTo ErrHandler

          Dim BucketingContents As Variant
          
2         If InStr(OutputFileTemplate, "??") = 0 Then
3             Throw "OutputFileTemplate must include '??' as a place holder for the bucket name or number"
4         End If
5         If VarType(BucketingFileOrArray) = vbString Then
6             BucketingContents = ThrowIfError(sFileShow(CStr(BucketingFileOrArray), ",", True, False, False, BucketingFileIsUnix))
7         ElseIf IsArray(BucketingFileOrArray) Then
8             BucketingContents = BucketingFileOrArray
9         End If
          
10        If sNCols(BucketingContents) <> 2 Then
11            Throw "BucketingFileOrArray must be the name of a comma-separated text file with two columns or must be an array with two columns"
12        End If

13        InputFileHeadersT = sArrayTranspose(ThrowIfError(sFileHeaders(InputFileName, ",", HeaderRowNumber)))
14        TargetFileMonikers = sVLookup(InputFileHeadersT, BucketingContents)
15        LookUpErrorText = sVLookup(1, sArrayRange(2, 3)) 'Avoid hard-coding "#Not found!"
16        ChooseVector = sArrayNot(sArrayEquals(TargetFileMonikers, LookUpErrorText))
17        If Not sAny(ChooseVector) Then
18            Throw "No column headers in the input file appear in the bucketing file's left column so no output files can be created"
19        End If

20        TargetFileMonikers = sMChoose(TargetFileMonikers, ChooseVector)
21        BucketsWrittenTo = sRemoveDuplicates(TargetFileMonikers, True)

          Dim MapCol1 As Variant 'Map to have three columns: 1) Index for reading input file, 2) Index indicating target file, 3) Index giving column number in target file
          Dim CounterArray() As Long
          Dim Map As Variant
          Dim MapCol2
          Dim MapCol3
          Dim NumColsRead As Long

22        MapCol1 = sMChoose(sIntegers(sNRows(InputFileHeadersT)), ChooseVector)
23        NumColsRead = sNRows(MapCol1)
24        TargetFileIndexes = sMatch(TargetFileMonikers, BucketsWrittenTo)
25        MapCol2 = TargetFileIndexes
26        MapCol3 = sReshape(0, NumColsRead, 1)
27        NumTargetFiles = sNRows(BucketsWrittenTo)

28        ReDim CounterArray(1 To NumTargetFiles)
29        If FirstColToAll Then
30            For i = 1 To NumTargetFiles
31                CounterArray(i) = 1 'Necessary since each target file has the first column of the input file written to it
32            Next
33        End If

34        For i = 1 To NumColsRead
35            CounterArray(MapCol2(i, 1)) = CounterArray(MapCol2(i, 1)) + 1
36            MapCol3(i, 1) = CounterArray(MapCol2(i, 1))
37        Next i

38        sFileSplitColumns = sArrayRange(MapCol1, MapCol2, MapCol3, TargetFileMonikers, BucketsWrittenTo)

39        Map = sArrayRange(MapCol1, MapCol2, MapCol3)

          Dim OutputLines() As Variant
          Dim ThisOutputLine() As String
40        ReDim OutputLines(1 To NumTargetFiles)

41        For i = 1 To NumTargetFiles
42            ReDim ThisOutputLine(1 To CounterArray(i)) 'CounterArray is left populated with the number of columns in each target file
43            OutputLines(i) = ThisOutputLine
44        Next i

45        ReDim TargetFileNames(1 To NumTargetFiles, 1 To 1)
46        ReDim TargetTextStreams(1 To NumTargetFiles)

47        Set FSO = New FileSystemObject

48        For i = 1 To NumTargetFiles
49            ThisBucket = BucketsWrittenTo(i, 1)
50            ThisFileName = Replace(OutputFileTemplate, "??", Format$(ThisBucket, BucketNumberFormat))
51            TargetFileNames(i, 1) = ThisFileName
52            Set TargetTextStreams(i) = FSO.OpenTextFile(ThisFileName, ForWriting, True)
53        Next i

54        Set t = FSO.OpenTextFile(InputFileName, ForReading, False)
55        For i = 1 To HeaderRowNumber - 1
56            t.SkipLine
57        Next

          Dim LineIn() As String
58        Do While Not t.atEndOfStream
59            LineIn = VBA.Split(t.ReadLine, ",")

60            If FirstColToAll Then
61                For i = 1 To NumTargetFiles 'write the first element of the input line to the first element of each output line
62                    OutputLines(i)(1) = LineIn(0)
63                Next i
64            End If

              'Read the map to allocate each element of the line read in to the appropiate element of the appropriate output line
65            For i = 1 To NumColsRead
66                OutputLines(Map(i, 2))(Map(i, 3)) = LineIn(Map(i, 1) - 1) 'The return from VBA.Split has Base 0, which accounts for the - 1 in this line of code
67            Next i
                
68            For i = 1 To NumTargetFiles
69                TargetTextStreams(i).Write VBA.Join(OutputLines(i), ",") + vbCrLf
70            Next i
71        Loop

72        For i = 1 To NumTargetFiles
73            TargetTextStreams(i).Close
74        Next i

75        sFileSplitColumns = TargetFileNames

76        Exit Function
ErrHandler:
77        sFileSplitColumns = "#sFileSplitColumns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderCopy
' Author    : Philip Swannell
' Purpose   : Copies a named folder (or folders) from one location to another, overwriting if the target
'             already exists. If the copy process fails, then an error string is returned.
' Arguments
' SourceFolder: Full name (with path) of the source folder. Can be an array, in which case TargetFolder
'             must be an array of the same dimensions. Does not matter if it has a
'             terminating backslash or not.
' TargetFolder: Full name (with path) of the target folder. Can be an array, in which case SourceFolder
'             must be an array of the same dimensions. Does not matter if it has a
'             terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderCopy(ByVal SourceFolder As Variant, ByVal TargetFolder As Variant)
Attribute sFolderCopy.VB_Description = "Copies a named folder (or folders) from one location to another, overwriting if the target already exists. If the copy process fails, then an error string is returned."
Attribute sFolderCopy.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFolderCopy = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFolderCopy = Broadcast2Args(FuncIdFolderCopy, SourceFolder, TargetFolder)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderExists
' Author    : Philip Swannell
' Purpose   : Returns TRUE if a folder of the given FolderPath exists on disk or FALSE otherwise.
' Arguments
' FolderPath: The full name of the folder. Does not matter if it has a terminating backslash or not.
'             This argument may be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderExists(FolderPath)
Attribute sFolderExists.VB_Description = "Returns TRUE if a folder of the given FolderPath exists on disk or FALSE otherwise."
Attribute sFolderExists.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFolderExists = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFolderExists = Broadcast1Arg(FuncIdFolderExists, FolderPath)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderIsWritable
' Author    : Philip Swannell
' Purpose   : Returns TRUE if a folder of the given FolderPath exists on disk and it is possible to
'             write new files to that FolderPath, returns FALSE otherwise.
' Arguments
' FolderPath: The full name of the folder. Does not matter if it has a terminating backslash or not.
'             This argument may be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderIsWritable(FolderPath)
Attribute sFolderIsWritable.VB_Description = "Returns TRUE if a folder of the given FolderPath exists on disk and it is possible to write new files to that FolderPath, returns FALSE otherwise."
Attribute sFolderIsWritable.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFolderIsWritable = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFolderIsWritable = Broadcast1Arg(FuncIdFolderIsWritable, FolderPath)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderMove
' Author    : Philip Swannell
' Purpose   : Moves a named folder (or folders) from one location to another, fails if ToFolder already
'             exists.
' Arguments
' FromFolder: Full name (with path) of the existing folder. Can be an array, in which case ToFolder must
'             be an array of the same dimensions. Does not matter if it has a terminating
'             backslash or not.
' ToFolder  : Full name (with path) of the new folder location. Can be an array, in which case
'             FromFolder must be an array of the same dimensions. Does not matter if it has
'             a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderMove(ByVal FromFolder As Variant, ByVal ToFolder As Variant)
Attribute sFolderMove.VB_Description = "Moves a named folder (or folders) from one location to another, fails if ToFolder already exists."
Attribute sFolderMove.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFolderMove = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFolderMove = Broadcast2Args(FuncIdFolderMove, FromFolder, ToFolder)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFormatDate
' Author    : Philip Swannell
' Date      : 30-Jan-2018
' Purpose   : Converts numbers to string representations of dates.
' Arguments
' DateNumbers: A number or an array of numbers. Non numeric entries in an array yield error strings in
'             the corresponding element of the output.
' DateFormat: The required Format, if omitted defaults to "dd-mmm-yyyy"
'
' Notes     : The function uses the same Date<->Number equivalence as Excel in which a Date of 0
'             represents 30-Dec-1899
'
'             Supported symbols are:
'             Date Symbols
'             Symbol    Range
'             d         1-31 (Day of month, with no leading zero).
'             dd        01-31 (Day of month, with a leading zero).
'             w         1-7 (Day of week, starting with Sunday = 1)
'             ww        1-53 (Week of year, with no leading zero; Week 1 starts on Jan 1).
'             mmm       Displays abbreviated month names (Hijri month names have no abbreviations).
'             mmmm      Displays full month names.
'             y         1-366 (Day of year)
'             yy        00-99 (Last two digits of year)
'             yyyy      100-9666 (Three- or Four-digit year)
'             Time Symbols
'             Symbol    Range
'             h         0-23 (1-12 with "Am" or "Pm" appended) (Hour of day, with no leading zero)
'             hh        0-23 (01-12 with "Am" or "Pm" appended) (Hour of day, with no leading zero)
'             n         0-59 (Minute of hour, with no leading zero)
'             nn        0-59 (Minute of hour, with a leading zero)
'             s         0-59 (Second of minute, with no leading zero)
'             ss        0-59 (Second of minute, with a leading zero)
'
'             The function is a "wrapper" to the VBA function Format. See
'             https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/Format-function-visual-basic-for-applications
' -----------------------------------------------------------------------------------------------------------------------
Function sFormatDate(ByVal DateNumbers, Optional DateFormat As String = "dd-mmm-yyyy")
Attribute sFormatDate.VB_Description = "Converts numbers to string representations of dates."
Attribute sFormatDate.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         If TypeName(DateNumbers) = "Range" Then DateNumbers = DateNumbers.Value2
3         Select Case NumDimensions(DateNumbers)
              Case 0
4                 Res = SafeFormat(DateNumbers, DateFormat)
5             Case 1
                  Dim i As Long
6                 Res = DateNumbers
7                 For i = LBound(Res) To UBound(Res)
8                     Res(i) = SafeFormat(Res(i), DateFormat)
9                 Next i
10            Case 2
                  Dim j As Long
11                Res = DateNumbers
12                For i = LBound(Res, 1) To UBound(Res, 1)
13                    For j = LBound(Res, 2) To UBound(Res, 2)
14                        Res(i, j) = SafeFormat(Res(i, j), DateFormat)
15                    Next j
16                Next i
17            Case Else
18                Throw "DateNumbers cannot have more than two dimensions"
19        End Select
20        sFormatDate = Res
21        Exit Function
ErrHandler:
22        sFormatDate = "#sFormatDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function SafeFormat(x, DateFormat As String) As String
1         On Error GoTo ErrHandler
2         If Not IsNumber(x) Then Throw "Input must be numeric"
3         SafeFormat = Format$(x, DateFormat)
4         Exit Function
ErrHandler:
5         SafeFormat = "#" & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sParseDate
' Author    : Philip Swannell
' Date      : 13-Dec-2017
' Purpose   : Convert a string representation of a date into the number Excel uses to represent that
'             date.
' Arguments
' DateStrings: A string or an array of strings.
' DateFormat: The date Format used such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read from
'             Windows regional settings. Repeated D's (or M's or Y's) are equivalent to
'             single instances, so that d-m-y and DD-MMM-YYYY are equivalent.
' -----------------------------------------------------------------------------------------------------------------------
Function sParseDate(DateStrings As Variant, Optional DateFormat As String, Optional ReturnFirstError As Boolean)
Attribute sParseDate.VB_Description = "Convert a string representation of a date into the number Excel uses to represent that date."
Attribute sParseDate.VB_ProcData.VB_Invoke_Func = " \n25"
          Dim DateSeparator As String
          Dim lDateFormat As Long

1         On Error GoTo ErrHandler

2         ParseDateFormat DateFormat, lDateFormat, DateSeparator

3         If VarType(DateStrings) = vbString Then
4             sParseDate = CoreParseDate(CStr(DateStrings), lDateFormat, DateSeparator, False)
5         Else
              Dim i As Long
              Dim j As Long
              Dim NC As Long
              Dim NR As Long
              Dim Result As Variant
6             Force2DArrayR DateStrings, NR, NC

7             Result = sReshape(0, NR, NC)
8             For i = 1 To NR
9                 For j = 1 To NC
10                    If ReturnFirstError Then
11                        Result(i, j) = ThrowIfError(CoreParseDate(CStr(DateStrings(i, j)), lDateFormat, DateSeparator, False))
12                    Else
13                        Result(i, j) = CoreParseDate(CStr(DateStrings(i, j)), lDateFormat, DateSeparator, False)
14                    End If
15                Next j
16            Next i
17            sParseDate = Result
18        End If
19        Exit Function
ErrHandler:
20        sParseDate = "#sParseDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub SpeedTestsFilesMerge()
          Dim i As Long
          Const NFiles = 20
          Const NR = 2750
          Const NC = 229
          Dim InputFiles
          Dim OutputFile As String

1         OutputFile = "c:\temp\medianFoo.txt"

2         InputFiles = sArrayConcatenate("c:\temp\Foo", sIntegers(NFiles), ".txt")
3         tic
4         For i = 1 To NFiles
5             ThrowIfError sFileSave(CStr(InputFiles(i, 1)), sReshape(i, NR, NC), ",")
6         Next
7         toc
8         tic
9         ThrowIfError sFilesMerge(InputFiles, OutputFile, ",", 0, 0)
10        toc
11        g sFileShow(OutputFile)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sSplitPath
' Author    : Philip Swannell
' Date      : 08-Sep-2013
' Purpose   : Splits FullFileName into either its file name (ReturnFileName = TRUE or omitted) or its
'             path (without terminating backslash, ReturnFileName = FALSE).
' Arguments
' FullFileName: The full name of the file, including path. Can be an array.
' ReturnFileName: TRUE to return the file name, FALSE to return the path.
' -----------------------------------------------------------------------------------------------------------------------
Function sSplitPath(ByVal FullFileName As Variant, Optional ReturnFileName As Boolean = True)
Attribute sSplitPath.VB_Description = "Splits FullFileName into either its file name (ReturnFileName = TRUE or omitted) or its path (without terminating backslash, ReturnFileName = FALSE)."
Attribute sSplitPath.VB_ProcData.VB_Invoke_Func = " \n26"
1         On Error GoTo ErrHandler
2         If VarType(FullFileName) = vbString Then
3             sSplitPath = CoreSplitPath(CStr(FullFileName), ReturnFileName)
4         Else
              Dim i As Long
              Dim j As Long
              Dim NC As Long
              Dim NR As Long
              Dim Res As Variant
5             Force2DArrayR FullFileName, NR, NC
6             Res = sReshape(vbNullString, NR, NC)
7             For i = 1 To NR
8                 For j = 1 To NC
9                     Res(i, j) = CoreSplitPath(FullFileName(i, j), ReturnFileName)
10                Next
11            Next
12            sSplitPath = Res
13        End If

14        Exit Function
ErrHandler:
15        sSplitPath = "#sSplitPath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sURLDownloadToFile
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Downloads the content of an internet URL location to a local file. If successful returns
'             TRUE.
' Arguments
' URLAddress: Internet URL address, such as "http://www.solum-financial.com/about-us/" Can be an array,
'             in which case FileName must be an array of the same size.
' FileName  : Name of a local file into which to store the downloaded contents, such as
'             "C:\temp\AboutUs.html" Can be an array, in which case URLAddress must be an
'             array of the same size.
' -----------------------------------------------------------------------------------------------------------------------
Function sURLDownloadToFile(ByVal URLAddress As Variant, ByVal FileName As Variant)
Attribute sURLDownloadToFile.VB_Description = "Downloads the content of an internet URL location to a local file. If successful, returns TRUE."
Attribute sURLDownloadToFile.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sURLDownloadToFile = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sURLDownloadToFile = Broadcast2Args(FuncIdURLDownloadToFile, URLAddress, FileName)
End Function

Private Function JoinPathVectorized(ByVal Paths1 As Variant, ByVal Paths2 As Variant)
          Dim c As Long
          Dim ColLock(1 To 2) As Boolean
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim R As Long
          Dim Result() As String
          Dim RowLock(1 To 2) As Boolean

1         On Error GoTo ErrHandler

2         NR = 1: NC = 1
3         Force2DArrayR Paths1, R, c
4         If R = 1 Then RowLock(1) = True
5         If c = 1 Then ColLock(1) = True
6         If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
7         If c > 1 Then If (NC = 1 Or c < NC) Then NC = c

8         Force2DArrayR Paths2, R, c
9         If R = 1 Then RowLock(2) = True
10        If c = 1 Then ColLock(2) = True
11        If R > 1 Then If (NR = 1 Or R < NR) Then NR = R
12        If c > 1 Then If (NC = 1 Or c < NC) Then NC = c

13        ReDim Result(1 To NR, 1 To NC)

14        For i = 1 To NR
15            For j = 1 To NC
16                Result(i, j) = CoreJoinPath(CStr(Paths1(IIf(RowLock(1), 1, i), IIf(ColLock(1), 1, j))), _
                      CStr(Paths2(IIf(RowLock(2), 1, i), IIf(ColLock(2), 1, j))))
17            Next j
18        Next i
                              
19        JoinPathVectorized = Result
20        Exit Function

21        Exit Function
ErrHandler:
22        Throw "#JoinPathVectorized (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sJoinPath
' Author    : Philip Swannell
' Date      : 14-Apr-2020
' Purpose   : Join path components into a full path. If some argument is an absolute path, then prior
'             components are dropped. Supports double dots to move up one level in the
'             folder heirarchy. All arguments may be arrays.
' If all inputs are singletons then return is singleton. If any input is an array (even 1-d array) then return is 2-d array.
' Arguments
' PathsToJoin:
'
' Notes     : Example:
'             The formula
'             sJoinPath("c:\foo", "baz\not this folder", "..", "this folder instead"
'             ,"myfile.txt")
'             returns
'             c:\foo\baz\this folder instead\myfile.txt
' -----------------------------------------------------------------------------------------------------------------------
Function sJoinPath(ParamArray PathsToJoin())
Attribute sJoinPath.VB_Description = "Join path components into a full path. If some argument is an absolute path, then prior components are dropped. Supports double dots to move up one level in the folder heirarchy. All arguments may be arrays."
Attribute sJoinPath.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim i As Long, LB As Long, UB As Long
          Dim AnyArrays As Boolean
          Dim Res As Variant
1         On Error GoTo ErrHandler
2         LB = LBound(PathsToJoin)
3         UB = UBound(PathsToJoin)

4         If LB = UB Then
5             sJoinPath = PathsToJoin(LB)
6             Exit Function
7         End If

8         For i = LB To UB
9             If VarType(PathsToJoin(i)) >= vbArray Then
10                AnyArrays = True
11                Exit For
12            End If
13        Next

          'What one really wants here is a "splat operator"
14        If Not AnyArrays Then
15            Select Case UB - LB + 1
                  Case 2
16                    sJoinPath = CoreJoinPath(PathsToJoin(LB), PathsToJoin(LB + 1))
17                    Exit Function
18                Case 3
19                    sJoinPath = CoreJoinPath(PathsToJoin(LB), PathsToJoin(LB + 1), PathsToJoin(LB + 2))
20                    Exit Function
21                Case 4
22                    sJoinPath = CoreJoinPath(PathsToJoin(LB), PathsToJoin(LB + 1), PathsToJoin(LB + 2), PathsToJoin(LB + 3))
23                    Exit Function
24                Case 5
25                    sJoinPath = CoreJoinPath(PathsToJoin(LB), PathsToJoin(LB + 1), PathsToJoin(LB + 2), PathsToJoin(LB + 3), PathsToJoin(LB + 4))
26                    Exit Function
27            End Select
28        End If

29        Res = JoinPathVectorized(PathsToJoin(LB), PathsToJoin(LB + 1))
30        For i = LB + 2 To UB
31            Res = JoinPathVectorized(Res, PathsToJoin(i))
32        Next

33        If Not AnyArrays Then
34            Res = Res(1, 1) 'will only hit this line if there are more than 5 arguments - unlikely in practice
35        End If
36        sJoinPath = Res

37        Exit Function
ErrHandler:
38        Throw "#sJoinPath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sRelativePath
' Author     : Philip Swannell
' Date       : 13-Mar-2022
' Purpose    : Returns RelativePath such that sJoinPath(BasePath,RelativePath) = FullPath
' -----------------------------------------------------------------------------------------------------------------------
Function sRelativePath(ByVal FullPath As Variant, ByVal BasePath As Variant)
Attribute sRelativePath.VB_Description = "The inverse of sJoinPath. For inputs FullPath and BasePath returns RelativePath such that sJoinPath(BasePath, RelativePath) = FullPath. Both arguments may be arrays."
Attribute sRelativePath.VB_ProcData.VB_Invoke_Func = " \n26"

1         On Error GoTo ErrHandler
2         If VarType(FullPath) < vbArray And VarType(BasePath) < vbArray Then
3             sRelativePath = CoreRelativePath(CStr(FullPath), CStr(BasePath))
4         Else
5             sRelativePath = Broadcast(FuncIdRelativePath, FullPath, BasePath)
6         End If

7         Exit Function
ErrHandler:
8         sRelativePath = "#sRelativePath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sWSLFileName
' Author    : Philip Swannell
' Date      : 16-Oct-2020
' Purpose   : Converts a file name to the name used to access it from Windows Subsystem for Linux.
' Arguments
' FileName: A file name or array of file names.
'
' Notes     : Example:
'             sWSLFileName("c:\temp\foo.txt") =
'             /mnt/c/temp/foo.txt
' -----------------------------------------------------------------------------------------------------------------------
Function sWSLFileName(ByVal FileName, Optional WindowsToLinux As Boolean = True)
Attribute sWSLFileName.VB_Description = "Converts file names between their representation on Windows and on WSL (Windows Subsystem for Linux)."
Attribute sWSLFileName.VB_ProcData.VB_Invoke_Func = " \n26"
1         On Error GoTo ErrHandler
2         If VarType(FileName) = vbString Then
3             sWSLFileName = CoreWSLFileName(CStr(FileName), WindowsToLinux)
4         Else
              Dim i As Long
              Dim j As Long
              Dim NC As Long
              Dim NR As Long
              Dim Res As Variant
5             Force2DArrayR FileName, NR, NC
6             Res = sReshape(vbNullString, NR, NC)
7             For i = 1 To NR
8                 For j = 1 To NC
9                     Res(i, j) = CoreWSLFileName(CStr(FileName(i, j)), WindowsToLinux)
10                Next
11            Next
12            sWSLFileName = Res
13        End If

14        Exit Function
ErrHandler:
15        sWSLFileName = "#sWSLFileName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function CoreWSLFileName(FileName As String, WindowsToLinux As Boolean)
          Dim Res As String

1         On Error GoTo ErrHandler
2         If WindowsToLinux Then
3             Res = Replace(FileName, "\", "/")
4             If Mid$(Res, 2, 1) = ":" Then
5                 Res = "/mnt/" & LCase(Left$(Res, 1)) & Mid$(Res, 3)
6             End If
7             CoreWSLFileName = Res
8         Else
9             If Left(FileName, 6) <> "\\wsl$" Then
10                CoreWSLFileName = "\\wsl$\Ubuntu-20.04" & Replace(FileName, "/", "\")
11            Else
12                CoreWSLFileName = Replace(FileName, "/", "\")
13            End If
14        End If

15        Exit Function
ErrHandler:
16        Throw "#CoreWSLFileName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
