Attribute VB_Name = "modCSVTest"
Option Explicit
Private Const TestFolder = "c:\temp\csvtest"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVSpeedTest
' Author     : Philip Swannell
' Date       : 19-Jul-2021
' Purpose    : Testing speed of sFileShow - record results below...
'2021-07-20 16:00:13.150   ====================================================================================================
'2021-07-20 16:00:13.152   SolumAddin version 2,166. Time of test = 20-Jul-2021 16:00:13
'2021-07-20 16:00:13.152   Time to read random doubles 10,000 rows, 100 cols = 3.68973119999998 seconds. File size = 18,180,900 bytes.
'2021-07-20 16:00:20.735   Time to read 10-char strings 10,000 rows, 100 cols = 3.03961860000004 seconds. File size = 11,010,000 bytes.
'2021-07-20 16:00:37.664   Time to read 10-char strings, all with line-feeds 10,000 rows, 100 cols = 11.2207711 seconds. File size = 14,010,000 bytes.
'2021-07-20 16:00:42.174   Time to read 1000 files = 2.80944360000001 seconds.
'2021-07-20 16:00:42.174   ====================================================================================================
'2021-07-20 19:47:32.791   ====================================================================================================
'2021-07-20 19:47:32.791   SolumAddin version 2,170. Time of test = 20-Jul-2021 19:47:32 Computer = PHILIP-LAPTOP
'2021-07-20 19:47:38.880   Time to read 1 file containing random doubles 10,000 rows, 100 cols = 2.04967829999987 seconds. File size = 18,180,900 bytes.
'2021-07-20 19:47:43.493   Time to read 1 file containing 10-char strings 10,000 rows, 100 cols = 1.96609960000023 seconds. File size = 11,010,000 bytes.
'2021-07-20 19:47:52.498   Time to read 1 file containing 10-char strings, all with line-feeds 10,000 rows, 100 cols = 5.65375040000072 seconds. File size = 14,010,000 bytes.
'2021-07-20 19:47:58.406   Time to write 1000 files = 4.86368840000068 seconds. Each file has 70 rows and 6 columns
'2021-07-20 19:48:06.752   Time to read 1000 files = 8.34504259999994 seconds. Each file has 70 rows and 6 columns
'2021-07-20 19:48:06.752   ====================================================================================================
'2021-07-27 11:15:51.958   ====================================================================================================
'2021-07-27 11:15:51.958   SolumAddin version 2,188. Time of test = 27-Jul-2021 11:15:51 Computer = PHILIP-LAPTOP
'2021-07-27 11:15:56.201   1.81936810002662 seconds to read 1 file containing random doubles 10,000 rows, 100 cols. File size = 18,180,900 bytes.
'2021-07-27 11:15:59.938   1.53920500003733 seconds to read 1 file containing UNquoted 10-char strings 10,000 rows, 100 cols. File size = 11,010,000 bytes.
'2021-07-27 11:16:08.306   5.71575269999448 seconds to read 1 file containing QUOTED 10-char strings 10,000 rows, 100 cols. File size = 13,010,000 bytes.
'2021-07-27 11:16:17.049   6.22214199998416 seconds to read 1 file containing 10-char strings, all with line-feeds 10,000 rows, 100 cols. File size = 15,010,000 bytes.
'2021-07-27 11:16:21.069   2.97768730006646 seconds to write 1000 files. Each file has 70 rows and 6 columns.
'2021-07-27 11:16:22.194   1.12387880007736 seconds to read 1000 files. Each file has 70 rows and 6 columns.
'2021-07-27 11:16:22.194   ====================================================================================================
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CSVSpeedTest()

          Dim FileName As String
          Dim NumRows As Long
          Dim NumCols As Long
          Dim Data, DataReread
          Dim t1 As Double, t2 As Double
          Dim OS As String

1         On Error GoTo ErrHandler

2         ShowFileInSnakeTail

3         NumRows = 10000
4         NumCols = 100
5         OS = "Windows"

6         ThrowIfError sCreateFolder(TestFolder)
7         MessageLogWrite String(100, "=")
8         MessageLogWrite "SolumAddin version " + CStr(Format$(sAddinVersionNumber(), "###,##0")) + ". Time of test = " + _
              Format$(Now, "dd-mmm-yyyy hh:mm:ss") + " Computer = " + Environ$("COMPUTERNAME")

          'Doubles only, cast back to doubles
9         sRandomSetSeed "MersenneTwister", 0
10        Data = sRandomVariable(NumRows, NumCols, , "MersenneTwister")
11        FileName = NameThatFile(OS, NumRows, NumCols, "Doubles", False, False)
12        ThrowIfError sFileSaveCSV(FileName, Data, , , , , OS)
13        t1 = sElapsedTime
14        DataReread = ThrowIfError(sFileShow(FileName, , True))
15        t2 = sElapsedTime
16        MessageLogWrite CStr(t2 - t1) + " seconds to read 1 file containing random doubles " + _
              Format$(NumRows, "###,##0") + " rows, " + Format$(NumCols, "###,##0") + " cols. " + _
              "File size = " + Format$(sFileInfo(FileName, "size"), "###,##0") + " bytes."
17        If Not sArraysNearlyIdentical(Data, DataReread) Then
18            MessageLogWrite "WARNING this did not roundtrip!"
19        End If

          '10-character strings, unquoted
20        Data = sReshape("abcdefghij", NumRows, NumCols)
21        FileName = NameThatFile(OS, NumRows, NumCols, "10-char-strings-unquoted", False, False)
22        ThrowIfError sFileSaveCSV(FileName, Data, False, , , , OS)
23        t1 = sElapsedTime
24        DataReread = ThrowIfError(sFileShow(FileName))
25        t2 = sElapsedTime
26        MessageLogWrite CStr(t2 - t1) + " seconds to read 1 file containing UNquoted 10-char strings " + _
              Format$(NumRows, "###,##0") + " rows, " + _
              Format$(NumCols, "###,##0") + " cols. File size = " + _
              Format$(sFileInfo(FileName, "size"), "###,##0") + " bytes."
27        If Not sArraysIdentical(Data, DataReread) Then
28            MessageLogWrite "WARNING this did not roundtrip!"
29        End If

          '10-character strings...
30        Data = sReshape("abcdefghij", NumRows, NumCols)
31        FileName = NameThatFile(OS, NumRows, NumCols, "10-char-strings", False, False)
32        ThrowIfError sFileSaveCSV(FileName, Data, , , , , OS)
33        t1 = sElapsedTime
34        DataReread = ThrowIfError(sFileShow(FileName))
35        t2 = sElapsedTime
36        MessageLogWrite CStr(t2 - t1) + " seconds to read 1 file containing QUOTED 10-char strings " + _
              Format$(NumRows, "###,##0") + " rows, " + _
              Format$(NumCols, "###,##0") + " cols. File size = " + _
              Format$(sFileInfo(FileName, "size"), "###,##0") + " bytes."
37        If Not sArraysIdentical(Data, DataReread) Then
38            MessageLogWrite "WARNING this did not roundtrip!"
39        End If

          '10-character strings ALL with linefeeds
40        Data = sReshape("abcd+" + vbCrLf + "efghi", NumRows, NumCols)
41        FileName = NameThatFile(OS, NumRows, NumCols, "10-char-strings-with-line-feeds", False, False)
42        ThrowIfError sFileSaveCSV(FileName, Data, , , , , OS)
43        t1 = sElapsedTime
44        DataReread = ThrowIfError(sFileShow(FileName, ","))
45        t2 = sElapsedTime
46        MessageLogWrite CStr(t2 - t1) + " seconds to read 1 file containing 10-char strings, all with line-feeds " + _
              Format$(NumRows, "###,##0") + " rows, " + Format$(NumCols, "###,##0") + " cols. File size = " + _
              Format$(sFileInfo(FileName, "size"), "###,##0") + " bytes."
47        If Not sArraysIdentical(Data, DataReread) Then
48            MessageLogWrite "WARNING this did not roundtrip!"
49        End If

          'Write and read many files
          Const NumFilesToRead = 1000
          Dim i As Long
          
          Const NumRowsSmall = 70
          Const NumColsSmall = 6
          Dim SmallFileName As String
          
          'Create files
50        t1 = sElapsedTime()
51        sRandomSetSeed "MersenneTwister", 100
52        For i = 1 To NumFilesToRead
53            SmallFileName = NameThatFile(OS, NumRowsSmall, NumColsSmall, Format$(i, "0000"), False, False)
54            Data = sRandomVariable(NumRowsSmall, NumColsSmall, , "MersenneTwister")
55            ThrowIfError sFileSave(SmallFileName, Data)
56        Next i
57        t2 = sElapsedTime()
58        MessageLogWrite CStr(t2 - t1) + " seconds to write " + CStr(NumFilesToRead) + " files. " + _
              "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
          
          'Read them back
59        t1 = sElapsedTime()
60        For i = 1 To NumFilesToRead
61            SmallFileName = NameThatFile(OS, NumRowsSmall, NumColsSmall, Format$(i, "0000"), False, False)
62            Data = ThrowIfError(sFileShow(SmallFileName, ",", True))
63        Next i
64        t2 = sElapsedTime()
65        MessageLogWrite CStr(t2 - t1) + " seconds to read " + CStr(NumFilesToRead) + " files. " + _
              "Each file has " + CStr(NumRowsSmall) + " rows and " + CStr(NumColsSmall) + " columns."
66        MessageLogWrite String(100, "=")

67        Exit Sub
ErrHandler:
68        SomethingWentWrong "#CSVSpeedTest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function NameThatFile(ByVal OS As String, NumRows As Long, NumCols As Long, ExtraInfo As String, Unicode As Boolean, Ragged As Boolean)
1         NameThatFile = sJoinPath(TestFolder, OS & "_" & Format$(NumRows, "0000") & "_x_" & Format$(NumCols, "000") & "_" & ExtraInfo & IIf(Unicode, "_Unicode", "_Ascii") & IIf(Ragged, "_Ragged", "_NotRagged") & ".csv")
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTestMulti
' Author     : Philip Swannell
' Date       : 22-Jul-2021
' Purpose    : Tests multiple times that sFileShow correctly round-trips data previously saved to disk by sFileSaveCSV.
'              Tests include:
'           *  Embedded line feeds in quoted strings.
'           *  Files with Windows, Unix or (old) Mac line endings.
'           *  Both unicode and ascii files.
'           *  Files with varying number of fields in each line (tricky since sFileSaveCSV does not support creating such files).
'           *  That the delimiter is automatically detected by sFileShow (reliable only if files have all strings quoted).
'           *  That unicode vs ascii is automatically detected.
'           *  That line endings are automatically detected.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CSVRoundTripTestMulti()
          Dim i As Long
          Dim AllowLineFeeds As Boolean
          Dim WholeFile As Boolean
          Dim EOL As String
          Dim Unicode As Boolean
          Dim Ragged As Boolean

1         On Error GoTo ErrHandler

2         ThrowIfError sCreateFolder(TestFolder)

3         For i = 1 To 200
4             If i Mod 10 = 0 Then
5                 Debug.Print i
6             End If
7             AllowLineFeeds = i Mod 2 = 0
8             WholeFile = i Mod 3 <> 0
9             Select Case i Mod 5
                  Case 0, 1
10                    EOL = "Windows"
11                Case 2, 3
12                    EOL = "Unix"
13                Case 4
14                    EOL = "Mac"
15            End Select
16            Select Case i Mod 7
                  Case 0, 1, 2, 3
17                    Unicode = False
18                Case 4, 5, 6
19                    Unicode = True
20            End Select
21            Select Case i Mod 11
                  Case 0, 1, 2, 3, 4, 5, 6, 7, 8
22                    Ragged = False
23                Case 9, 10
24                    Ragged = True
25            End Select

26            If Not CSVRoundTripTest(AllowLineFeeds, EOL, WholeFile, Unicode, Ragged) Then
                  'No need to throw since CSVRoundTripTest calls Stop
                  '  Throw ("OhOh")
27            End If
28        Next i

29        Exit Sub
ErrHandler:
30        SomethingWentWrong "#CSVRoundTripTestMulti (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CSVRoundTripTest
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : Do one round-trip test, i.e. generate random data, save it to file with sFileSaveCSV, read back the data
'              with sFileShow and then test if the data has correctly round-tripped.
' Parameters :
'  AllowLineFeed: Do we permit line feed characters inside strings?
'  EOL          : Windows, Unix or Mac line endings? When reading back we detect automatically.
'  WholeFile    : Do we read back the whole file, or only part of it via the StartRow -> NumCols arguments?
'  Unicode      : True = strings are Unicode, False = strings are ascii.
'  Ragged       : Make the lines in the saved files have varying line lengths. This is a bit tricky we have to a) amend
'                 the random data generated to include Empty values for a random number of elements at the right of each
'                 row (function MakeArrayRagged) write the data to file and then call function MakeCSVRagged that uses
'                 regular expressions to remove repeated commas from each line of the file.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CSVRoundTripTest(AllowLineFeed As Boolean, EOL As String, WholeFile As Boolean, Unicode As Boolean, Ragged As Boolean)
          Dim Data1 As Variant, Data2 As Variant
          Dim FileName As String
          Const DateFormat = "yyyy-mmm-dd"
          Dim identical As Boolean
          Dim StartRow As Long, StartCol As Long, NumRows As Long, NumCols As Long, ExtraInfo As String

1         On Error GoTo ErrHandler

2         Data1 = RandomArray(AllowLineFeed, Unicode)
3         If Ragged Then Data1 = MakeArrayRagged(Data1)

4         If AllowLineFeed Then
5             ExtraInfo = "Random_with_LF"
6         Else
7             ExtraInfo = "Random_No_LF"
8         End If

9         FileName = NameThatFile(EOL, sNRows(Data1), sNCols(Data1), ExtraInfo, Unicode, Ragged)

          'For round-tripping to be always possible we quote all strings when writing
10        ThrowIfError sFileSaveCSV(FileName, Data1, True, DateFormat, , Unicode, EOL)
11        If Ragged Then MakeCSVRagged FileName, EOL

12        If Not WholeFile Then
13            StartRow = Rnd() * sNRows(Data1)
14            If StartRow = 0 Then StartRow = 1
15            StartCol = Rnd() * sNCols(Data1)
16            If StartCol = 0 Then StartCol = 1

17            NumRows = Rnd() * (sNRows(Data1) - StartRow + 1)
18            NumCols = Rnd() * (sNCols(Data1) - StartCol + 1)
19        Else
20            StartRow = 1
21            StartCol = 1
22            NumRows = 0
23            NumCols = 0
24        End If

          'Note automatically detecting EOL character and delimiter in call to sFileShow.
25        Data2 = sFileShow(FileName, , True, "Date", True, , DateFormat, , , , StartRow, StartCol, NumRows, NumCols, , Empty)
          'Test the test
          'Data2(1, 1) = "This is different"
          
26        If sIsErrorString(Data2) Then
27            Stop
28        End If

29        If Not WholeFile Then
30            Data1 = sSubArray(Data1, StartRow, StartCol, NumRows, NumCols)
31        End If

32        identical = sArraysIdentical(Data1, Data2)
          'with Ragged files when we are reading only part of the file there is a circumstance _
           when Data1 and Data2 can legitimately differ. Its the case where the data read back _
           has fewer columns (but same num rows), but the corresponding sub-part of the original data was all empty. _
           This happens because sFileShow, when reading only some lines of a ragged file does not know the length of the lines it did not read.
33        If Not identical Then
34            If Ragged Then
35                If Not WholeFile Then
36                    If sNRows(Data1) = sNRows(Data2) Then
37                        If sNCols(Data2) < sNCols(Data1) Then
                              Dim AmendedData2
38                            AmendedData2 = sArrayRange(Data2, sReshape(Empty, sNRows(Data2), sNCols(Data1) - sNCols(Data2)))
39                            If sArraysIdentical(Data1, AmendedData2) Then
40                                identical = True
41                            End If
42                        End If
43                    End If
44                End If
45            End If
46        End If
47        If Not identical Then
              Dim Spacer
48            Spacer = sReshape(String(20, "="), 1, sNCols(Data1))
49            g sArrayStack(Data1, Spacer, Data2, Spacer, sDiffTwoArrays(Data1, Data2))
50        End If

51        CSVRoundTripTest = identical
52        If Not identical Then Stop

53        Exit Function
ErrHandler:
54        Throw "#CSVRoundTripTest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomArray(AllowLineFeed As Boolean, Unicode As Boolean)
          Dim NRows As Long, NCols As Long, i As Long, j As Long
          Const DateFormat = "yyyy-mmm-dd"
          Const MAXROWS = 50
          Const MAXCOLS = 5
          
          Dim Res() As Variant

1         On Error GoTo ErrHandler
2         NRows = 1 + Rnd * MAXROWS
3         NCols = 1 + Rnd() * MAXCOLS
4         ReDim Res(1 To NRows, 1 To NCols)

5         For i = 1 To NRows
6             For j = 1 To NCols
7                 Res(i, j) = RandomVariant(DateFormat, AllowLineFeed, Unicode)
8             Next j
9         Next i
10        If AllowLineFeed Then
11            Res(1, 1) = "this" & vbLf & "definitely" & vbCr & "has" & vbCrLf & "line feeds"
12        End If

13        RandomArray = Res

14        Exit Function
ErrHandler:
15        Throw "#RandomArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomVariant(DateFormat As String, AllowLineFeed As Boolean, Unicode As Boolean)

          Dim x
          Const NUMTYPES = 10

1         On Error GoTo ErrHandler
2         x = Rnd()

3         If x < 1 / NUMTYPES Then
              'Bool
4             RandomVariant = Rnd() < 0.5
5         ElseIf x < 2 / NUMTYPES Then
              'Long
6             RandomVariant = CLng((Rnd() - 0.5) * 2000000)
              'Double, casting to String and back to Double yields a double with an _
               exact string representation, which avoids failing the round trip
7             RandomVariant = CDbl(CStr((Rnd() - 0.5) * 2 * 10 ^ ((Rnd() - 0.5) * 20)))
8         ElseIf x < 3 / NUMTYPES Then

9         ElseIf x < 4 / NUMTYPES Then
              'String
10            RandomVariant = RandomString(AllowLineFeed, Unicode)
11        ElseIf x < 5 / NUMTYPES Then
              'Date
12            RandomVariant = RandomDate()
13        ElseIf x < 6 / NUMTYPES Then
              'Empty string
14            RandomVariant = ""
15        ElseIf x < 7 / NUMTYPES Then
              'String that looks like a number
16            RandomVariant = CStr(CLng((Rnd() - 0.5) * 2000000))
17        ElseIf x < 8 / NUMTYPES Then
              'String that looks like a date
18            RandomVariant = sFormatDate(CLng(RandomDate()), DateFormat)
19        ElseIf x < 9 / NUMTYPES Then
              'String that looks like Boolean
20            RandomVariant = CStr(Rnd() < 0.5)
21        ElseIf x < 10 / NUMTYPES Then
              'Empty
22            RandomVariant = Empty
23        End If

24        Exit Function
ErrHandler:
25        Throw "#RandomVariant (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RandomString
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : A random string either unicode or ascii, max length 20
' -----------------------------------------------------------------------------------------------------------------------
Function RandomString(AllowLineFeed As Boolean, Unicode As Boolean)
          Dim length As Long
          Dim i As Long
          Dim Res As String
          Const MAXLEN = 20
1         On Error GoTo ErrHandler
2         length = CLng(1 + Rnd() * MAXLEN)
3         Res = String(length, " ")

4         For i = 1 To length
5             If Unicode Then
6                 Mid(Res, i, 1) = ChrW(33 + Rnd() * 370)
7             Else
8                 Mid(Res, i, 1) = Chr(34 + Rnd() * 88)
9             End If

10            If Not AllowLineFeed Then
11                If Mid(Res, i, 1) = vbLf Or Mid(Res, i, 1) = vbCr Then
12                    Mid(Res, i, 1) = " "
13                End If
14            End If
15        Next
16        RandomString = Res

17        Exit Function
ErrHandler:
18        Throw "#RandomString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomDate()
1         On Error GoTo ErrHandler
2         RandomDate = CDate(CLng(Rnd() * Date * 2))
3         Exit Function
ErrHandler:
4         Throw "#RandomDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeArrayRagged
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : For each row of an array choose random number n less than number of cols and make the n right most columns empty
'              also guarantee that one row will not have an empty right most column.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MakeArrayRagged(Data)

          Dim NR As Long, NC As Long
          Dim i As Long, j As Long
          Dim RowToLeaveUnchanged As Long

1         On Error GoTo ErrHandler
2         NR = sNRows(Data)
3         NC = sNCols(Data)
4         RowToLeaveUnchanged = 1 + Rnd() * (NR - 1)

5         For i = 1 To NR
6             If i = RowToLeaveUnchanged Then
7                 If IsEmpty(Data(i, NC)) Then
8                     Data(i, NC) = "Not empty!"
9                 End If
10            Else
11                For j = CLng(1 + Rnd() * (NC - 1)) To NC
12                    Data(i, j) = Empty
13                Next
14            End If
15        Next
16        MakeArrayRagged = Data

17        Exit Function
ErrHandler:
18        Throw "#MakeArrayRagged (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeCSVRagged
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : To help test sFileShow's behaviour with ragged files, takes a file and removes consecutive commas at line ends
' -----------------------------------------------------------------------------------------------------------------------
Private Function MakeCSVRagged(FileName As String, EOL As String)

          Dim Contents As String
          Dim NewContents As String
          Dim Unicode As Boolean
          Dim FSO As New Scripting.FileSystemObject
          Dim t As Scripting.TextStream
          Dim RegularExpression As String

1         On Error GoTo ErrHandler
2         Unicode = IsUnicodeFile(FileName)
3         EOL = vbCrLf

4         Set t = FSO.GetFile(FileName).OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

5         Contents = t.ReadAll
6         t.Close

7         Select Case EOL
              Case "Windows", vbCrLf
8                 RegularExpression = ",+\r\n"
9             Case "Unix", vbLf
10                RegularExpression = ",+\r\n"
11            Case "Mac", vbCr
12                RegularExpression = ",+\r"
13            Case Else
14                Throw "EOL not recognised"
15        End Select

16        NewContents = sRegExReplace(Contents, RegularExpression, EOL)

17        Set t = FSO.GetFile(FileName).OpenAsTextStream(ForWriting, IIf(Unicode, TristateTrue, TristateFalse))
18        t.Write NewContents
19        t.Close

20        Exit Function
ErrHandler:
21        Throw "#MakeCSVRagged (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
