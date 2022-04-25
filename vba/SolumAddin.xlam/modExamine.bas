Attribute VB_Name = "modExamine"
Option Explicit

Enum EnmExaminationMethod
    ExMthdSpreadsheet = 1
    ExMthdDebugWindow = 2
    ExMthdTextEditor = 3
    ExMthdLogFile = 4
End Enum

Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "USER32" (ByVal lHwnd As LongPtr, ByVal lCmdShow As Long) As Boolean
Private Const GW_HWNDNEXT = 2

'Adapted from https://stackoverflow.com/questions/25098263/how-to-use-findwindow-to-find-a-visible-or-invisible-window-with-a-partial-name
Private Function GetHandleFromPartialCaption(ByRef lwnd As LongPtr, ByVal sCaption As String) As Boolean

          Dim lhWndP As LongPtr
          Dim sStr As String
1         GetHandleFromPartialCaption = False
2         lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
3         Do While lhWndP <> 0
4             sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
5             GetWindowText lhWndP, sStr, Len(sStr)
6             sStr = Left$(sStr, Len(sStr) - 1)
7             If InStr(1, sStr, sCaption) > 0 Then
8                 GetHandleFromPartialCaption = True
9                 lwnd = lhWndP
10                Exit Do
11            End If
12            lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
13        Loop

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : g
' Author    : Philip Swannell
' Date      : 15-May-2015
' Purpose   : Debugging routine. Prints an array, singleton. Dictionary or Collection
'             to the to a worksheet, a text editor or the debug window
' -----------------------------------------------------------------------------------------------------------------------
Sub g(TheData, Optional Method As EnmExaminationMethod = ExMthdTextEditor)
1         On Error GoTo ErrHandler
2         Select Case Method
              Case ExMthdSpreadsheet
3                 ExamineInSheet TheData
4             Case Else
5                 Examine TheData, Method
6         End Select
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#g (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExamineInSheet
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Debugging routine. Displays the contents of an array variable in a newly
'             created workbook workbook is then saved to to a sub-folder of the temporary folder and any other workbooks
'             from that directory are closed.
'             26-Jan-2017
'             Now also copes with Dictionary and Collection - displayed in JSON Format
' -----------------------------------------------------------------------------------------------------------------------
Sub ExamineInSheet(ByVal Data)
          Dim R As Range
          Dim wb As Excel.Workbook
          Dim PathToSaveData As String

1         On Error GoTo ErrHandler

2         PathToSaveData = Environ$("Temp") & "\@" & gCompanyName & "Temp"
3         If Not sFolderExists(PathToSaveData) Then
4             sCreateFolder PathToSaveData
5         End If

6         If TypeName(Data) = "Dictionary" Or TypeName(Data) = "Collection" Then
7             Data = ConvertToJson(Data, 3, AS_RowByRow)
8             Data = sTokeniseString(CStr(Data), vbCrLf)
9         End If

10        Force2DArray Data

          'Avoid excel's annoying habit of interpreting strings as something else...
11        Data = sArrayExcelString(Data)
          'close the other ones to avoid masses of open workbooks
12        For Each wb In Application.Workbooks
13            If InStr(LCase$(wb.Path), LCase$(PathToSaveData)) > 0 Then
14                wb.Close True
15            End If
16        Next

17        Set wb = Application.Workbooks.Add
18        Set R = wb.Worksheets(1).Range("A1").Resize(UBound(Data, 1) - LBound(Data, 1) + 1, UBound(Data, 2) - LBound(Data, 2) + 1)
19        R.Value = Data
20        R.Columns.AutoFit
21        wb.SaveAs PathToSaveData + "\DataToExamine_" + Format$(Now, "yyyy-mm-dd_hh-mm-ss")

22        Exit Sub
ErrHandler:
23        Throw "#ExamineInSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CleanTemporaryDirectory
' Author    : Philip Swannell
' Date      : 18-May-2015
' Purpose   : Deletes files more than 10 days old in subdirectory SolumDebug of temporary directory
'             Such files are created by methods Examine, TemporaryMessage and MessageLogWrite
' PGS 21-March-2021 Amended to not delete the MessageLog files, turns out can be helpful to refer back to see what I did
'             on a given day...
' -----------------------------------------------------------------------------------------------------------------------
Sub CleanTemporaryDirectory()
          Dim FolderName As String
          Dim FSO As Scripting.FileSystemObject
          Dim myFile As Scripting.file
          Dim myFolder As Scripting.Folder
          Dim NumDaysToKeep As Long

1         On Error GoTo ErrHandler
2         FolderName = Environ$("Temp")
3         If Right$(FolderName, 1) <> "\" Then FolderName = FolderName + "\"
4         FolderName = FolderName + "@" + gCompanyName + "Temp"
5         If sEquals(True, sFolderExists(FolderName)) Then

6             Set FSO = New Scripting.FileSystemObject
7             Set myFolder = FSO.GetFolder(FolderName)
8             For Each myFile In myFolder.Files
9                 If InStr(myFile.Name, "MessageLog") > 0 Then
10                    NumDaysToKeep = 1000000
11                Else
12                    NumDaysToKeep = 4
13                End If
14                If myFile.DateCreated < Date - NumDaysToKeep Then
15                    myFile.Delete
16                End If
17            Next
18        End If

19        Exit Sub
ErrHandler:
20        Throw "#CleanTemporaryDirectory (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExamineDictionary
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Sub of Examine since code for dictionaries has little in common with that for arrays...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ExamineDictionary(dct As Variant, Method As EnmExaminationMethod)

          Dim Representation As String
          Dim TopText As String
          Const NumSeparators = 20
          Dim FileName As String
          Dim FSO As Scripting.FileSystemObject
          Dim TS As Scripting.TextStream

1         On Error GoTo ErrHandler
2         TopText = String(NumSeparators, "=") + vbCrLf + _
              "TypeName = " & TypeName(dct) + vbCrLf + _
              "JSON representation:" + vbCrLf + _
              String(NumSeparators, "=")

3         Representation = ConvertToJson(dct, 3, AS_RowByRow)

4         Select Case Method
              Case ExMthdDebugWindow
5                 Debug.Print TopText
6                 Debug.Print Representation
7             Case ExMthdTextEditor
8                 Set FSO = New FileSystemObject
                  Dim Folder As String
9                 Folder = Environ$("Temp")
10                If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
11                Folder = Folder + "@" + gCompanyName + "Temp"
12                sCreateFolder Folder
13                FileName = Folder + "\DataToExamine_" + Format$(Now, "yyyy-mm-dd_hh-mm-ss") + "-" + Format$(FileNameCounter(), "000") + ".txt"
14                Set TS = FSO.CreateTextFile(FileName, True)
15                TS.Write TopText + vbCrLf
16                TS.WriteLine Representation
17                TS.Write String(NumSeparators, "=")
18                TS.Close
19                ShowFileInTextEditor FileName
20            Case Else
21                Throw "Unrecognised examination method"
22        End Select
23        Exit Sub
ErrHandler:
24        Throw "#ExamineDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileNameCounter
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Returns the "next" file number to use as part of name of temp files
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileNameCounter()
          Static c As Long
1         c = (c + 1) Mod 1000
2         FileNameCounter = c
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Examine
' Author    : Philip Swannell
' Date      : 15-May-2015
' Purpose   : Alternative debugging routine. g prints to a worksheet but wont work when debugging a
'             call to a worksheet function.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Examine(ByVal TheData As Variant, Optional Method As EnmExaminationMethod)
          Dim i As Long
          Dim j As Long
          Dim MaxWidths As Variant
          Dim NC As Long
          Dim ND As Long
          Dim NR As Long
          Dim origNC As Long
          Dim origNR As Long
          Dim TheType As String
          Const NExtraSpaces As Long = 2
          Dim CarriageReturnsFound As Boolean
          Dim FileName As String
          Dim FSO As Scripting.FileSystemObject
          Dim LineFeedsFound As Boolean
          Dim LongStrings
          Dim NumSeparators As Long
          Dim StartPoints
          Dim StringsFound As Boolean
          Dim DatesFound As Boolean
          Dim TopText As String
          Dim TS As Scripting.TextStream
          Const EscapeCr = "<CR>"        '"&#13;"
          Const EscapeLf = "<LF>"        '"&#10;"
          Dim LeftNumbers As Variant
          Dim OneDLB As Long
          Dim OneDUB As Long
          Dim TopNumbers As Variant

1         On Error GoTo ErrHandler

2         If Method = ExMthdTextEditor Then CleanTemporaryDirectory
3         TheType = TypeName(TheData)
4         If TheType = "Dictionary" Or TheType = "Collection" Then
5             ExamineDictionary TheData, Method
6             Exit Sub
7         End If

8         If TheType = "Range" Then TheData = TheData.Value2

9         ND = NumDimensions(TheData)
10        If ND = 1 Then
11            OneDLB = LBound(TheData)
12            OneDUB = UBound(TheData)
13        End If

14        Force2DArray TheData, origNR, origNC

15        Select Case ND
              Case 0
16                LeftNumbers = vbNullString
17                TopNumbers = vbNullString
18            Case 1
19                TheData = sArrayTranspose(TheData)
20                TopNumbers = vbNullString
21                LeftNumbers = sIntegers(OneDUB - OneDLB + 1)
22                If OneDLB <> 1 Then
23                    LeftNumbers = sArrayAdd(LeftNumbers, OneDLB - 1)
24                End If
25                origNR = origNC
26            Case 2
27                LeftNumbers = sIntegers(origNR)
28                If LBound(TheData, 1) <> 1 Then
29                    LeftNumbers = sArrayAdd(LeftNumbers, LBound(TheData, 1) - 1)
30                End If
31                TopNumbers = sArrayTranspose(sIntegers(origNC))
32                If LBound(TheData, 2) <> 1 Then
33                    TopNumbers = sArrayAdd(TopNumbers, LBound(TheData, 2) - 1)
34                End If
35        End Select

36        If ND > 0 Then
37            TheData = sArrayRange(LeftNumbers, sReshape("|", origNR, 1), TheData)
38            TheData = sArrayStack(sArrayRange(vbNullString, "|", TopNumbers), TheData)
39        End If

40        NC = sNCols(TheData)
41        NR = sNRows(TheData)
42        MaxWidths = sReshape(0, 1, NC)

43        For j = 1 To NC
44            MaxWidths(1, j) = 0
45            For i = 1 To NR
46                If VarType(TheData(i, j)) <> vbString Then
47                    If VarType(TheData(i, j)) = vbDate Then DatesFound = True
48                    TheData(i, j) = NonStringToString(TheData(i, j))
49                Else
50                    If j > 2 Then
51                        If i > 1 Then
52                            StringsFound = True
53                            If InStr(TheData(i, j), vbCr) > 0 Then
54                                CarriageReturnsFound = True
55                                TheData(i, j) = Replace(TheData(i, j), vbCr, EscapeCr)
56                            End If
57                            If InStr(TheData(i, j), vbLf) > 0 Then
58                                LineFeedsFound = True
59                                TheData(i, j) = Replace(TheData(i, j), vbLf, EscapeLf)
60                            End If
61                            TheData(i, j) = "'" + (TheData(i, j)) + "'"
62                        End If
63                    End If
64                End If
65                If Len(TheData(i, j)) > MaxWidths(1, j) Then
66                    MaxWidths(1, j) = Len(TheData(i, j))
67                End If
68            Next i
69        Next j

70        LongStrings = sReshape(String(sRowSum(MaxWidths)(1, 1) + (NC - 1) * NExtraSpaces, " "), NR, 1)
71        StartPoints = sReshape(0, 1, NC)
72        StartPoints(1, 1) = 1
73        For j = 2 To NC
74            StartPoints(1, j) = StartPoints(1, j - 1) + MaxWidths(1, j - 1) + NExtraSpaces
75        Next j

76        For i = 1 To NR
77            For j = 1 To NC
78                If Len(TheData(i, j)) > 0 Then
79                    Mid$(LongStrings(i, 1), StartPoints(1, j), Len(TheData(i, j))) = TheData(i, j)
80                End If
81            Next j
82        Next i

83        NumSeparators = Len(LongStrings(1, 1))
84        If NumSeparators < 20 Then NumSeparators = 20
85        TopText = String(NumSeparators, "=") + vbCrLf + _
              "TypeName = " & TheType + vbCrLf + _
              "Num Dimensions = " & CStr(ND)
86        If TheType = "String" Then
87            TopText = TopText + vbCrLf + _
                  "Length = " & Format$(Len(TheData(1, 1)), "###,##0")
88        End If
89        If ND = 2 Then
90            TopText = TopText + vbCrLf + _
                  "Num Rows = " & CStr(origNR) + " (" + CStr(LeftNumbers(1, 1)) + " to " + CStr(LeftNumbers(sNRows(LeftNumbers), 1)) + ")" + vbCrLf + _
                  "Num Cols = " & CStr(origNC) + " (" + CStr(TopNumbers(1, 1)) + " to " + CStr(TopNumbers(1, sNCols(TopNumbers))) + ")"
91        ElseIf ND = 1 Then
92            TopText = TopText + vbCrLf + _
                  "Num elements = " & CStr(origNC) + " (" + CStr(OneDLB) + " to " + CStr(OneDUB) + ")"
93        End If
94        If StringsFound Then
95            TopText = TopText + vbCrLf + _
                  "Strings are displayed surrounded by single quote characters."
96        End If
97        If DatesFound Then
98            TopText = TopText + vbCrLf + _
                  "Some elements are dates, formatted either 'dd-mmm-yyy' or 'dd-mmm-yyyy hh:mm:ss' as appropriate."
99        End If

100       If CarriageReturnsFound Or LineFeedsFound Then
101           TopText = TopText + vbCrLf + _
                  "Carriage return characters and/or line feed characters were encountered. They are displayed as " + EscapeCr + " and " + EscapeLf + " respectively."
102       End If

103       If ND > 0 Then
104           TopText = TopText + vbCrLf + _
                  String(NumSeparators, "-")
105       End If

106       If Method = ExMthdDebugWindow Then
107           Debug.Print TopText
108           For i = 1 To NR
109               If i = 2 Then Debug.Print String(NumSeparators, "-")
110               Debug.Print LongStrings(i, 1)
111           Next
112           Debug.Print String(NumSeparators, "=")
113       ElseIf Method = ExMthdTextEditor Then
114           Set FSO = New FileSystemObject
              Dim Folder As String
115           Folder = Environ$("Temp")
116           If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
117           Folder = Folder + "@" + gCompanyName + "Temp"
118           sCreateFolder Folder
119           FileName = Folder + "\DataToExamine_" + Format$(Now, "yyyy-mm-dd_hh-mm-ss") + "-" + Format$(FileNameCounter, "000") + ".txt"
120           Set TS = FSO.CreateTextFile(FileName, True, Unicode:=True)
121           TS.Write TopText + vbCrLf
122           For i = 1 To NR
123               If i = 2 Then TS.WriteLine String(NumSeparators, "-")
124               TS.WriteLine LongStrings(i, 1)
125           Next
126           TS.WriteLine String(NumSeparators, "=")
127           TS.Close
128           ShowFileInTextEditor FileName
129       ElseIf Method = ExMthdLogFile Then
130           FileName = MessagesLogFileName()        'creates a file if it doesn't already exist
              Const NumSp = 27        'Changing this? Then also change constant of the same name in method MessageLogWrite
131           Set FSO = New FileSystemObject
132           Set TS = FSO.OpenTextFile(FileName, ForAppending)
133           TS.WriteLine Format$(Now, "dd-mmm-yyyy hh:mm:ss") + "   " + "Data written to this log file via call to " & gAddinName & ".modExamine.Examine."
134           TopText = String(NumSp, " ") + Replace(TopText, vbCrLf, vbCrLf + String(NumSp, " "))
135           TS.Write TopText + vbCrLf
136           For i = 1 To NR
137               If i = 2 Then TS.WriteLine String(NumSp, " ") + String(NumSeparators, "-")
138               TS.WriteLine String(NumSp, " ") + LongStrings(i, 1)
139           Next
140           TS.WriteLine String(NumSp, " ") + String(NumSeparators, "=")
141           TS.Close
142       Else
143           Throw "Unrecognised examination method"
144       End If
145       Exit Sub
ErrHandler:
146       Throw "#Examine (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowFileInTextEditor
' Author    : Philip Swannell
' Date      : 05-Nov-2015
' Purpose   : Displayes the contents of a text file in NotePad++ or NotePad. Used by
'             methods Examine and ShowMessages.
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowFileInTextEditor(FileName As String, Optional ScrollToEnd As Boolean)
          Const NotePadPP = "C:\Program Files (x86)\Notepad++\notepad++.exe"        'Location if one follows the defaults in the Notepad++ installer on 64 bit Wndows
          Const NotePadPP_B = "C:\Program Files\Notepad++\notepad++.exe"        'Location if one follows the defaults in the Notepad++ installer on 32 bit Windows
          Const NotePad = "C:\Windows\System32\notepad.exe"        'Standard location.
          Dim NotepadPPOptions As String
          Const DQ = """"

1         On Error GoTo ErrHandler

2         If ScrollToEnd Then
3             NotepadPPOptions = " -alwaysOnTop -nosession -n100000000 "        'This is a bit hacky, since I'm not sure if there's a way to say scroll to last line
4         Else
5             NotepadPPOptions = " -alwaysOnTop -nosession "
6         End If

7         If sFileExists(NotePadPP) Then
8             Shell DQ + NotePadPP + DQ + NotepadPPOptions + DQ + FileName + DQ, vbNormalFocus
9         ElseIf sFileExists(NotePadPP_B) Then
10            Shell DQ + NotePadPP_B + DQ + NotepadPPOptions + DQ + FileName + DQ, vbNormalFocus
11        ElseIf sFileExists(NotePad) Then
12            Shell DQ + NotePad + DQ + " " + DQ + FileName + DQ, vbNormalFocus
13        Else
14            Throw "Cannot find Notepad++ or Notepad executable files. Locations searched:" + vbLf + NotePadPP + vbLf + NotePadPP_B + vbLf + NotePad
15        End If
16        Exit Sub
ErrHandler:
17        Throw "#ShowFileInTextEditor (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowFileInSnakeTail
' Author    : Philip Swannell
' Date      : 03-Oct-2018
' Purpose   : Pops up SnakeTail displaying a file
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowFileInSnakeTail(Optional ByVal FileName As String, Optional ThrowIfNoExe As Boolean = True)
          Const DQ = """"
          Const SnakeTail = "C:\Program Files\SnakeTail\SnakeTail.exe"
          Const SnakeTail_B = "C:\Program Files (x86)\SnakeTail\SnakeTail.exe"
          Dim caption
          Dim H As LongPtr
          
1         On Error GoTo ErrHandler

2         If FileName = vbNullString Then FileName = MessagesLogFileName()

          'May be open already? - Code below is bugged if we are trying to open a file of the same name, but different folder to an existing file
          'Also leaves the snaketail window minimised if that's its current state...
          Dim EN As Long
3         caption = "SnakeTail - [" & sSplitPath(FileName)
4         GetHandleFromPartialCaption H, caption
5         If H <> 0 Then
              'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
6             ShowWindow H, 9
7         End If

8         On Error Resume Next
9         AppActivate caption
10        EN = Err.Number
11        On Error GoTo ErrHandler
12        If EN = 0 Then

13            Exit Sub
14        End If
15        If sFileExists(SnakeTail) Then
16            Shell DQ + SnakeTail + DQ + " " + DQ + FileName + DQ, vbNormalFocus
17        ElseIf sFileExists(SnakeTail_B) Then
18            Shell DQ + SnakeTail_B + DQ + " " + DQ + FileName + DQ, vbNormalFocus
19        ElseIf ThrowIfNoExe Then
20            Throw "Cannot find SnakeTail executable at locations:" + vbLf + vbLf + SnakeTail + vbLf + SnakeTail_B + vbLf + vbLf + "Please install SnakeTail from http://snakenest.com/snaketail/"
21        End If
22        Exit Sub
ErrHandler:
23        Throw "#ShowFileInSnakeTail (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NonStringToString
' Author    : Philip Swannell
' Date      : 29-May-2015
' Purpose   : Convert non-string to string in a way that mimics how the non-string would
'             be displayed in an Excel cell. Used by functions such as ConcatenateStrings
'             and Examine (aka g)
' -----------------------------------------------------------------------------------------------------------------------
Function NonStringToString(x As Variant, Optional AddSingleQuotesToStings As Boolean = False)
          Dim Res As String
1         On Error GoTo ErrHandler
2         If IsError(x) Then
3             Select Case CStr(x)
                  Case "Error 2007"
4                     Res = "#DIV/0!"
5                 Case "Error 2029"
6                     Res = "#NAME?"
7                 Case "Error 2023"
8                     Res = "#REF!"
9                 Case "Error 2036"
10                    Res = "#NUM!"
11                Case "Error 2000"
12                    Res = "#NULL!"
13                Case "Error 2042"
14                    Res = "#N/A"
15                Case "Error 2015"
16                    Res = "#VALUE!"
17                Case "Error 2045"
18                    Res = "#SPILL!"
19                Case "Error 2047"
20                    Res = "#BLOCKED!"
21                Case "Error 2046"
22                    Res = "#CONNECT!"
23                Case "Error 2048"
24                    Res = "#UNKNOWN!"
25                Case "Error 2043"
26                    Res = "#GETTING_DATA!"
27                Case Else
28                    Res = CStr(x)        'should never hit this line...
29            End Select
30        ElseIf VarType(x) = vbDate Then
31            If CDbl(x) = CLng(x) Then
32                Res = Format$(x, "dd-mmm-yyyy")
33            Else
34                Res = Format$(x, "dd-mmm-yyyy hh:mm:ss")
35            End If
36        ElseIf IsNull(x) Then
37            Res = "null" 'Follow how json represents Null as lower-case null
38        ElseIf VarType(x) = vbString And AddSingleQuotesToStings Then
39            Res = "'" + x + "'"
40        Else
41            Res = SafeCStr(x)        'Converts Empty to null string. Prior to 15 Jan 2017 Empty was converted to "Empty"
42        End If
43        NonStringToString = Res
44        Exit Function
ErrHandler:
45        Throw "#NonStringToString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
