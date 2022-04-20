Attribute VB_Name = "modDirList"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modDirList
' Author    : Philip Swannell
' Date      : 14-June-2013
' Purpose   : Function to return a list of all files in a folder.
'             This function re-written in May 2019 to use old-fashioned-but-faster Dir$
'             function rather than FileSystemObject. Speedup can be quite large - factor
'             of 56 in a recent use case of mine - 16,207 files in a single folder.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_FSO As Scripting.FileSystemObject

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestsDirList
' Author     : Philip Swannell
' Date       : 17-May-2019
' Purpose    : Test sDirList versus sDirListFSO
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestsDirList()

          Dim ColumnTemplate As String
          Dim FileFilter As String
          Dim FilesOrFolders As String
          Dim Folder As String
          Dim FolderFilter As String
          Dim H As Long
          Dim i As Long
          Dim identical As Boolean
          Dim j As Long
          Dim k As Long
          Dim M As Long
          Dim Recurse As Boolean
          Dim Res1 As Variant
          Dim Res2 As Variant
          Dim WithHeaders As Boolean

1         On Error GoTo ErrHandler
2         Folder = "c:\TestsDirList"
3         WithHeaders = False
4         ColumnTemplate = "TNF^SMCA#"

5         For H = 1 To 2
6             Recurse = Choose(H, True, False)
7             For i = 1 To 3
8                 FilesOrFolders = Choose(i, "F", "D", "FD")
9                 For j = 1 To 3
10                    FileFilter = Choose(j, "*", vbNullString, "*B*")
11                    For k = 1 To 2
12                        M = M + 1
13                        FolderFilter = Choose(k, "*", vbNullString)
14                        Res1 = sDirList(Folder, Recurse, WithHeaders, ColumnTemplate, FilesOrFolders, FileFilter, FolderFilter)
15                        Res2 = sDirListFSO(Folder, Recurse, WithHeaders, ColumnTemplate, FilesOrFolders, FileFilter, FolderFilter)
16                        If sIsErrorString(Res1) And sIsErrorString(Res2) Then
17                            identical = True
18                        Else
19                            identical = sArraysIdentical(Res1, Res2)
20                        End If
21                        Debug.Print M, "Identical = ", identical, "Recurse = ", Recurse, "FilesOrFolders = ", FilesOrFolders, "FileFilter = ", FileFilter, "FolderFilter = ", FolderFilter
22                    Next k
23                Next j
24            Next i
25        Next H
26        Exit Sub
ErrHandler:
27        SomethingWentWrong "#TestsDirList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub TestSpeedsDirList()
          Dim Folder As String
          Dim Res1 As Variant
          Dim Res2 As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
1         On Error GoTo ErrHandler
          '16,207 files
2         Folder = "\\solumsbs\Philip Shared\ISDA SIMM\2019\2019-04-30-CTOnly\Concentration Threshold\EQ Delta single name\TurnOver data"
3         t1 = sElapsedTime()
4         Res1 = sDirList(Folder, False, False, "F", "F")
5         t2 = sElapsedTime()
6         Res2 = sDirListFSO(Folder, False, False, "F", "F")
7         t3 = sElapsedTime()

8         Debug.Print "Identical results:", sArraysIdentical(Res1, Res2), "sDirListFSO:", t3 - t2, "sDirList:", t2 - t1, "Ratio:", (t3 - t2) / (t2 - t1)
          'Identical results:          True          sDirListFSO:   114.980655099964           sDirList:      2.02470259997062           Ratio:         56.7889106783544
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#TestSpeedsDirList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDirList
' Author    : Philip Swannell
' Date      : 16-May-2019
' Purpose   : Lists files or folders in a given folder(s). Can recurse into sub-folders; can filter on
'             file/folder names, and the return can be sorted. Use ColumnTemplate to
'             specify the columns of the return.
' Arguments
' Folder    : Path to folder e.g. "c:\temp" or "c:\temp\". May be an array to list files in disparate
'             folders.
' Recurse   : Logical value to determine whether or not the returned array includes files from sub
'             folders. Defaults to False for no recursion into sub-folders.
' WithHeaders: Logical value to determine whether the returned array has a header row. Defaults to False
'             for no headers.
' ColumnTemplate: Columns: N = Name, F = FullName, R = RelativeName, D = Folder, S = Size, # = MD5, M =
'             DateLastModified, A = DateLastAccessed, C = DateCreated, T = File or folDer,
'             B = Attributes, v or ^ = sort on prior column. Defaults to F if Recurse is
'             True, N otherwise
' FilesOrFolders: F to list files, D to list folders or FD to list both. Defaults to F.
' FileFilter: Optional case-insensitive filter on file names. Either:
'             a) Wild-cards * = multiple characters, ? = single character e.g. *.txt to
'             match text files
'             b) RegExp followed by a regular expression e.g.  RegExp\.dll$|\.txt$ to match
'             both dll files and text files
' FolderFilter: Optional case-insensitive filter on folder names, with the same syntax as FileFilter. A
'             file is returned if its name matches the FileFilter and its folder matches
'             the FolderFilter.
' UnicodeSupport: FALSE (the default) for good speed but poor unicode support - unicode in file names become
'             'similar' ascii or question marks; recursion into unicode folders does not
'             work. TRUE for full unicode support, but the function is slower in this case.
'
' Notes     : Example:
'             =sDirList("C:\Temp\",TRUE,TRUE,"FSv")
'             returns a two-column array, with a header row, of the full names and sizes of
'             all files in c:\Temp and is subfolders. The return is sorted by file size in
'             descending order, i.e. largest file at the top.
'
'             When using FilePattern against folders with a large number (many thousands)
'             of files then a "simple" pattern using * for multiple characters and ? for a
'             single character will be faster than using a regular expression. This is
'             because the low-level VB function Dir supports pattern matching but does not
'             support regular expresion matching. There is no equivalent speed advantage in
'             the case of the FolderFilter.
' -----------------------------------------------------------------------------------------------------------------------
Function sDirList(ByVal Folder As Variant, Optional ByVal Recurse As Boolean, Optional WithHeaders As Boolean, _
        Optional ByVal ColumnTemplate As String, Optional FilesOrFolders As String = "F", _
        Optional ByVal FileFilter As String, Optional ByVal FolderFilter As String, Optional UnicodeSupport As Boolean = False)
Attribute sDirList.VB_Description = "Lists files or folders in a given folder(s). Can recurse into sub-folders; can filter on file/folder names, and the return can be sorted. Use ColumnTemplate to specify the columns of the return."
Attribute sDirList.VB_ProcData.VB_Invoke_Func = " \n26"
          
          Dim Ascending As Boolean
          Dim AttributesToSearchFor As Long
          Dim CurrentFolder As String
          Dim D As Scripting.Folder
          Dim DoesFileMatch As Boolean
          Dim DoesFolderMatch As Boolean
          Dim DoesSubFolderMatch As Boolean
          Dim DoFiles As Boolean
          Dim DoFiltering As Boolean
          Dim DoFolders As Boolean
          Dim DoSort As Boolean
          Dim F As Scripting.file
          Dim FileName As String
          Dim FilePattern As String
          Dim FileRegExp As VBScript_RegExp_55.RegExp
          Dim FirstArgToDir As String
          Dim FolderCollection As Collection
          Dim FolderRegExp As VBScript_RegExp_55.RegExp
          Dim Hasher As Object
          Dim i As Long
          Dim isFolder As Boolean
          Dim NeedFSO As Boolean
          Dim OneLine() As Variant
          Dim origFolder As String
          Dim Result As Variant
          Dim SortColNum As Long
          Dim STK As clsStacker
          Dim ZeroByteHash As String
          
1         On Error GoTo ErrHandler

2         If UnicodeSupport Then
3             sDirList = sDirListFSO(Folder, Recurse, WithHeaders, ColumnTemplate, FilesOrFolders, FileFilter, FolderFilter)
4             Exit Function
5         End If

6         Set m_FSO = New Scripting.FileSystemObject

7         If FunctionWizardActive() Then
8             sDirList = "#Disabled in Function Dialog!"
9             Exit Function
10        End If

11        ParseFilesOrFolders FilesOrFolders, DoFiles, DoFolders

12        If DoFiles Then
13            If InStr(ColumnTemplate, "#") > 0 Then
14                Set Hasher = CreateHasher(True, ZeroByteHash)
15            End If
16        End If

17        ParseColumnTemplate ColumnTemplate, OneLine, DoSort, SortColNum, Ascending, Recurse

18        Set STK = CreateStacker()
19        If WithHeaders Then
20            STK.Stack2D OneLine
21        End If
          
22        ParseFilters FileFilter, FolderFilter, DoFiltering, FilePattern, FileRegExp, FolderRegExp
23        If Not DoFiltering Then
24            DoesFileMatch = True
25            DoesFolderMatch = True
26            DoesSubFolderMatch = True
27        End If
          
28        Set FolderCollection = InitialiseFolderCollection(Folder, origFolder)
          
          'This section of the code provides a speedup. We are a) Recursing; and b) Doing filtering of files using a _
           FilePattern - the kind of filtering that is supported by the low-level Dir function. But the Dir function _
           is inconvenient - we need (assuming no folder filtering) ALL of the sub-folders and the only way to get _
           them is to iterate without use of a pattern since Dir doesn't greatly distinguish between folders and _
           files. Thus the (big) speed up of being able to pass a pattern to Dir is lost. _
           The solution is to iterate to find the sub-folders using the FileSystemObject (coded in method sFolderList), _
           add those sub-folders to the FoldersCollection and then switch Recurse to False. In this way we can continue _
           to use the FilePattern for speed when calling Dir.
29        If Recurse And DoFiles And DoFolders And DoFiltering And FilePattern <> vbNullString Then
              Dim DoFolderFiltering As Boolean
              Dim FolderCollection2 As Collection
              Dim STK2 As clsStacker
              Dim SubFolders As Variant
              'Can't just write Set FolderCollection2 = FolderCollection as that would be a reference, not a copy
30            Set FolderCollection2 = CopyCollection(FolderCollection)
31            Set STK2 = CreateStacker()
32            DoFolderFiltering = FolderRegExp.Pattern <> vbNullString
33            sFolderList FolderCollection2, Recurse, STK2, "F", DoFolderFiltering, FolderRegExp, vbNullString
34            If STK2.NumRows > 0 Then
35                SubFolders = STK2.ReportInTranspose
36                For i = 1 To STK2.NumRows
37                    FolderCollection.Add SubFolders(1, i) & "\"
38                Next i
39            End If
40            Recurse = False
41        End If
          
          'If we are listing folders but not files then we use the file system object (as coded in sFolderList) _
           since when using Dir it's not possible to iterate over Folders only, instead you have to iterate over _
           folders and files (For Dir, a folder is a file albeit with a different Attribute value). So the speed _
           advantage of Dir is likely to be more than offset by the fact that there are typically many more files than folders.
42        If DoFolders And Not DoFiles Then
43            sFolderList FolderCollection, Recurse, STK, ColumnTemplate, DoFiltering, FolderRegExp, origFolder
44        Else
45            If DoFiles Then
46                AttributesToSearchFor = vbNormal + vbReadOnly + vbHidden + vbSystem
47            End If
48            If Recurse Or DoFolders Then
49                AttributesToSearchFor = AttributesToSearchFor + vbDirectory
50            End If

51            NeedFSO = InStr(ColumnTemplate, "C") > 0 Or _
                  InStr(ColumnTemplate, "A") > 0 Or _
                  (InStr(ColumnTemplate, "S") > 0 And DoFolders)

52            Do While FolderCollection.Count > 0
53                CurrentFolder = FolderCollection.item(1)
54                FolderCollection.Remove 1
55                If DoFiltering Then
56                    DoesFolderMatch = FolderRegExp.Test(Left$(CurrentFolder, Len(CurrentFolder) - 1))
57                End If

58                If Recurse Or DoFolders Then 'Because in both these cases we want folders to be returned by the call to Dir
59                    FirstArgToDir = CurrentFolder
60                Else
61                    FirstArgToDir = CurrentFolder & FilePattern
62                End If
                    
63                FileName = DirWrap(FirstArgToDir, AttributesToSearchFor)
64                Do While Len(FileName) > 0
65                    If (FileName <> ".") And (FileName <> "..") Then 'Strange property of Dir that the first two items it returns are "." and ".."
66                        If Not (Recurse Or DoFolders) Then
67                            isFolder = False 'No need to call GetAttr in this case
68                        Else
69                            isFolder = TestIfFolder(CurrentFolder & FileName)
70                        End If
71                        If isFolder Then
72                            If Recurse Then
73                                FolderCollection.Add CurrentFolder & FileName & "\"
74                            End If
75                            If DoFiltering Then
76                                DoesSubFolderMatch = FolderRegExp.Test(CurrentFolder & FileName)
77                            End If
78                            If DoFolders And DoesSubFolderMatch Then
79                                If NeedFSO Then
80                                    Set D = m_FSO.GetFolder(CurrentFolder & FileName)
81                                End If
82                                PopulateOneLine OneLine, FileName, Left$(CurrentFolder, Len(CurrentFolder) - 1), vbNullString, ColumnTemplate, origFolder, True, D
83                                STK.Stack2D OneLine
84                            End If
85                        Else
86                            If DoFiltering Then
87                                DoesFileMatch = FileRegExp.Test(FileName)
88                            End If
89                            If DoFiles And DoesFolderMatch And DoesFileMatch Then
90                                If NeedFSO Then
91                                    Set F = m_FSO.GetFile(CurrentFolder & FileName)
92                                End If
93                                PopulateOneLine OneLine, FileName, Left$(CurrentFolder, Len(CurrentFolder) - 1), vbNullString, ColumnTemplate, origFolder, False, F, Hasher
94                                STK.Stack2D OneLine
95                            End If
96                        End If
97                    End If
98                    FileName = DirWrap()
99                Loop
100           Loop
101       End If
102       Result = STK.Report

103       If STK.NumRows > 1 Then
104           If DoSort Then
105               If WithHeaders Then
106                   Result = sArrayStack(sSubArray(Result, 1, 1, 1), sSortedArray(sSubArray(Result, 2), SortColNum, , , Ascending))
107               Else
108                   Result = sSortedArray(Result, SortColNum, , , Ascending)
109               End If
110           End If
111       ElseIf STK.NumRows = 0 Then
112           If DoFiles And DoFolders Then
113               Result = "#No files or folders found!"
114           ElseIf DoFiles Then
115               Result = "#No files found!"
116           Else
117               Result = "#No folders found!"
118           End If
119       End If

120       sDirList = Result

121       Exit Function
ErrHandler:
122       sDirList = "#sDirList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub wesfvsdfvsd()
1         TestIfFolder "D:\Philip\OneDrive\iTunes Media\Music\Janácek, Leoš (1854 - 1928)"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestIfFolder
' Author     : Philip Swannell
' Date       : 18-May-2019
' Purpose    : GetAttr has nasty habit of raising "Permission Denied" error e.g. on c:\PageFle.sys, so wrap
' -----------------------------------------------------------------------------------------------------------------------
Private Function TestIfFolder(PathName As String) As Boolean
          Dim EN As Long
          Dim file As Scripting.file
          Dim fld As Scripting.Folder

1         On Error Resume Next
2         TestIfFolder = GetAttr(PathName) And vbDirectory
3         If Err.Number = 0 Then
4             Exit Function
5         End If

6         On Error GoTo ErrHandler
7         If m_FSO Is Nothing Then
8             Set m_FSO = New Scripting.FileSystemObject
9         End If

10        On Error Resume Next
11        Set fld = m_FSO.GetFolder(PathName)
12        EN = Err.Number
13        On Error GoTo ErrHandler
14        If EN = 0 Then
15            TestIfFolder = True
16            Exit Function
17        End If

18        On Error Resume Next
19        Set file = m_FSO.GetFile(PathName)
20        EN = Err.Number
21        On Error GoTo ErrHandler
22        If EN = 0 Then
23            TestIfFolder = False
24            Exit Function
25        End If

26        If Len(PathName) > 255 Then
27            Throw "Encountered file paths over 255 characters. Try setting argument UnicodeSupport to TRUE (though that makes function slower)"
28        Else
29            Throw "Cannot determine if file '" + PathName + "' is a File or a Folder. Possible cause: the file name includes Unicode characters. Try setting argument UnicodeSupport to TRUE (though that makes function slower)"
30        End If

31        Exit Function
ErrHandler:
32        Throw "#TestIfFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitFullName
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Splits a FullFileName into the pair of Path and FileName, returned by reference
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SplitFullName(FullFileName As String, ByRef Path As String, ByRef FileName As String)
1         On Error GoTo ErrHandler
          Dim SlashPos As Long
2         SlashPos = InStrRev(FullFileName, "\")
3         If SlashPos = 0 Then Throw "Cannot split path '" + FullFileName + "'"
4         FileName = Mid$(FullFileName, SlashPos + 1)
5         Path = Left$(FullFileName, SlashPos - 1)
6         Exit Sub
ErrHandler:
7         Throw "#SplitFullName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PopulateOneLine
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Encapsulation...
' Parameters :
'  OneLine       : A 1-row array to be populated by reference
'  FileName, FolderName, FullName: Provide EITHER FileName + FolderName OR FullName
' -----------------------------------------------------------------------------------------------------------------------
Private Sub PopulateOneLine(ByRef OneLine, ByVal FileName As String, ByVal FolderName As String, ByVal FullName As String, _
        ColumnTemplate As String, origFolder As String, isFolder As Boolean, o As Object, Optional Hasher As Object, Optional ZeroByteHash As String)
          Dim i As Long
1         On Error GoTo ErrHandler

2         If Len(FullName) > 0 Then
3             SplitFullName FullName, FolderName, FileName
4         Else
5             FullName = FolderName + "\" + FileName
6         End If

7         For i = 1 To Len(ColumnTemplate)
8             Select Case UCase$(Mid$(ColumnTemplate, i, 1))
                  Case "N"
9                     OneLine(1, i) = FileName
10                Case "F"
11                    OneLine(1, i) = FullName
12                Case "D"
13                    OneLine(1, i) = FolderName
14                Case "T"
15                    OneLine(1, i) = IIf(isFolder, "D", "F")
16                Case "B"
17                    On Error Resume Next
18                    OneLine(1, i) = GetAttr(FullName)
19                    If Err.Number <> 0 Then
20                        OneLine(1, i) = "#" + Err.Description + "!"
21                    End If
22                    On Error GoTo ErrHandler
23                Case "R"
24                    OneLine(1, i) = RelativeFileName(FullName, origFolder)
25                Case "S"
26                    If isFolder Then
27                        On Error Resume Next
28                        OneLine(1, i) = o.Size
29                        If Err.Number <> 0 Then
30                            OneLine(1, i) = "#" + Err.Description + "!"
31                        End If
32                        On Error GoTo ErrHandler
33                    Else
34                        On Error Resume Next
35                        OneLine(1, i) = FileLen(FullName)
36                        If Err.Number <> 0 Then
37                            OneLine(1, i) = "#" + Err.Description + "!"
38                        ElseIf OneLine(1, i) < 0 Then 'FileLen returns a Long and file sizes between 2^31 and 2^32 are shown as negative, so correct
39                            OneLine(1, i) = 4294967296# + OneLine(1, i)
40                        End If
41                        On Error GoTo ErrHandler
42                    End If
43                Case "M"
44                    OneLine(1, i) = CDbl(FileDateTime(FullName))
45                Case "C"
46                    OneLine(1, i) = CDbl(o.DateCreated)
47                Case "A"
48                    OneLine(1, i) = CDbl(o.DateLastAccessed)
49                Case "#"
50                    If isFolder Then
51                        OneLine(1, i) = vbNullString
52                    Else
                          Dim TheHash As String
53                        HashFromFileName Hasher, FullName, TheHash, ZeroByteHash
54                        OneLine(1, i) = TheHash
55                    End If
56                Case Else
57                    Throw "Unexpected character in ColumnTemplate" 'Should not hit this line since we test the ColumnTemplate earlier
58            End Select
59        Next i

60        Exit Sub
ErrHandler:
61        Throw "#PopulateOneLine (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseFilesOrFolders
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Parse the FilesOrFolders argument to sDirList, setting a couple of Booleans
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseFilesOrFolders(FilesOrFolders As String, ByRef DoFiles As Boolean, DoFolders As Boolean)
1         Select Case UCase$(FilesOrFolders)
              Case "F"
2                 DoFiles = True
3                 DoFolders = False
4             Case "D"
5                 DoFiles = False
6                 DoFolders = True
7             Case "FD", "DF"
8                 DoFiles = True
9                 DoFolders = True
10            Case Else
11                Throw "FilesOrFolders must be 'F' (or omitted) or 'D' or 'FD'"
12        End Select
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DirWrap
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Dir$ fails to give the PathName in its error description, which is quite annoying, so wrap.
' -----------------------------------------------------------------------------------------------------------------------
Private Function DirWrap(Optional PathName As String, Optional Attributes) As String
          Dim CopyOfErr As String
          Static PriorPathName As String
1         On Error GoTo ErrHandler
2         If PathName <> vbNullString Then
3             PriorPathName = PathName
4             DirWrap = Dir$(PathName, Attributes)
5         Else
6             DirWrap = Dir$() ', Attributes)
7         End If
8         Exit Function
ErrHandler:
9         CopyOfErr = Err.Description
10        If PathName <> vbNullString Then
11            CopyOfErr = "Error calling Dir$ - " + CopyOfErr + " PathName = '" + PathName + "'"
12        ElseIf PriorPathName <> vbNullString Then
13            CopyOfErr = "Error calling Dir$ - " + CopyOfErr + " PriorPathName = '" + PriorPathName + "'"
14        End If
15        Throw "#DirWrap (line " & CStr(Erl) + "): " & CopyOfErr & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CopyCollection
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Creates a copy of a collection, rather than a reference.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CopyCollection(c As Collection) As Collection
          Dim i As Variant
1         On Error GoTo ErrHandler
2         Set CopyCollection = New Collection
3         For Each i In c
4             CopyCollection.Add i
5         Next i
6         Exit Function
ErrHandler:
7         Throw "#CopyCollection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sFolderList
' Author     : Philip Swannell
' Date       : 13-May-2019
' Purpose    : In the event that DirList is iterating over Folders, but not Files (FilesOrFolders = "D")
'              then it's likely to be quicker to use the FileSystemObject.
'              That's because "Calling Dir with the vbDirectory attribute does not continually return subdirectories"
'              https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dir-function
' -----------------------------------------------------------------------------------------------------------------------
Private Sub sFolderList(ByVal FolderCollection As Collection, ByVal Recurse As Boolean, ByRef STK As clsStacker, _
        ColumnTemplate2 As String, ByVal DoFiltering As Boolean, FolderRegExp As VBScript_RegExp_55.RegExp, origFolder As String)
          
          Dim Child As Scripting.Folder
          Dim DoThis As Boolean
          Dim OneLine() As Variant
          Dim Parent As Scripting.Folder
          
1         On Error GoTo ErrHandler
2         DoThis = True
3         ReDim OneLine(1 To 1, 1 To Len(ColumnTemplate2))

4         Do While FolderCollection.Count > 0
5             Set Parent = m_FSO.GetFolder(FolderCollection.item(1))
6             FolderCollection.Remove 1
7             For Each Child In Parent.SubFolders
8                 If DoFiltering Then
9                     DoThis = FolderRegExp.Test(Child.Path)
10                End If
11                If DoThis Then
12                    PopulateOneLine OneLine, vbNullString, vbNullString, Child.Path, ColumnTemplate2, origFolder, True, Child
13                    STK.Stack2D OneLine
14                End If
15                If Recurse Then
16                    FolderCollection.Add Child.Path
17                End If
18            Next Child
19        Loop

20        Exit Sub
ErrHandler:
21        Throw "#sFolderList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseFilters
' Author     : Philip Swannell
' Date       : 14-May-2019
' Purpose    : Parse the pair of passed-in arguments FileFilter and FolderFilter, generating a FilePattern (that can be
'              interpreted by Dir) or alternatively a FileRegExp and also a FolderRegExp (no FolderPattern since we never
'              use a pattern when searching for folders)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseFilters(ByVal FileFilter As String, ByVal FolderFilter As String, _
        ByRef DoFiltering As Boolean, _
        ByRef FilePattern As String, _
        ByRef FileRegExp As VBScript_RegExp_55.RegExp, ByRef FolderRegExp As VBScript_RegExp_55.RegExp)
          
1         On Error GoTo ErrHandler
          
2         If FileFilter = vbNullString And FolderFilter = vbNullString Then
3             DoFiltering = False
4             Exit Sub
5         Else
6             DoFiltering = True
7         End If

8         Set FileRegExp = New RegExp
9         If LCase$(Left$(FileFilter, 6)) = "regexp" Then
10            If VarType(sIsRegMatch(Mid$(FileFilter, 7), "Foo", False)) <> vbBoolean Then Throw "That FileFilter is invalid because the regular expression '" + Mid$(FileFilter, 7) + "' has incorrect syntax"
11            FileRegExp.Pattern = Mid$(FileFilter, 7)
12            FileRegExp.IgnoreCase = True
13            FileRegExp.Global = False
14            FilePattern = vbNullString
15        ElseIf FileFilter <> vbNullString Then
              'In this case we give the FileRegExp the same "Power" as the FilePattern
16            FileRegExp.Pattern = RegExpFromPattern(FileFilter)
17            FileRegExp.IgnoreCase = True
18            FileRegExp.Global = False
19            FilePattern = FileFilter
20        End If

21        Set FolderRegExp = New RegExp
22        If LCase$(Left$(FolderFilter, 6)) = "regexp" Then
23            If VarType(sIsRegMatch(Mid$(FolderFilter, 7), "Foo", False)) <> vbBoolean Then Throw "That FolderFilter is invalid because the regular expression '" + Mid$(FolderFilter, 7) + "' has incorrect syntax"
24            FolderRegExp.Pattern = Mid$(FolderFilter, 7)
25            FolderRegExp.IgnoreCase = True
26            FolderRegExp.Global = False
27        ElseIf FolderFilter <> vbNullString Then
28            FolderRegExp.Pattern = RegExpFromPattern(FolderFilter)
29            FolderRegExp.IgnoreCase = True
30            FolderRegExp.Global = False
31        End If

32        Exit Sub
ErrHandler:
33        Throw "#ParseFilters (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function RelativeFileName(FullName As String, origFolder As String)
1         If LCase$(origFolder) = LCase$(Left$(FullName, Len(origFolder))) Then
2             RelativeFileName = Mid$(FullName, Len(origFolder) + 1)
3         Else
4             RelativeFileName = FullName
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InitialiseFolderCollection
' Author     : Philip Swannell
' Date       : 13-May-2019
' Purpose    : Returns a collection containing the strings in Folder with backslashes appended if necessary and with corrected capitalisation to match folders on disk.
'              origFolder is a ByRef argument required by the caller for "Relative addresses"
' -----------------------------------------------------------------------------------------------------------------------
Private Function InitialiseFolderCollection(ByVal Folder As Variant, ByRef origFolder As String) As Collection

          Dim cln As Collection
          Dim cln2 As Collection
          Dim F As Scripting.Folder

1         On Error GoTo ErrHandler
2         Set cln = New Collection

3         If VarType(Folder) = vbString Then
4             origFolder = Folder
5             cln.Add Folder
6         Else
              Dim j As Long
              Dim k As Long
              Dim NC As Long
              Dim NR As Long
7             Force2DArrayR Folder, NR, NC
8             If NR = 1 Then
9                 If NC = 1 Then
10                    origFolder = CStr(Folder(1, 1))
11                End If
12            End If
13            For j = 1 To NR
14                For k = 1 To NC
15                    If VarType(Folder(j, k)) = vbString Then
16                        cln.Add Folder(j, k)
17                    End If
18                Next
19            Next
20        End If

21        Set cln2 = New Collection
          Dim Str
22        For Each Str In cln
23            Set F = Nothing
24            On Error Resume Next
25            Set F = m_FSO.GetFolder(Str)
26            On Error GoTo ErrHandler
27            If F Is Nothing Then Throw "Cannot find folder '" + Str + "'"

28            cln2.Add F.Path + IIf(Right$(F.Path, 1) = "\", vbNullString, "\") 'Ensures correct capitalisation
29        Next
          
30        If cln2.Count = 0 Then
31            Throw "No folders specified"
32        End If

33        Set InitialiseFolderCollection = cln2

34        Exit Function
ErrHandler:
35        Throw "#InitialiseFolderCollection (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseColumnTemplate
' Author     : Philip Swannell
' Date       : 13-May-2019
' Purpose    : Strip the ColumnTemplate of the "^" and "v" characters that indicate sorting, populate by reference variables.
'              Headers is populated even if the return from sDirList will not have headers that validates the ColumnTemplate
' -----------------------------------------------------------------------------------------------------------------------
Sub ParseColumnTemplate(ByRef ColumnTemplate As String, ByRef Headers As Variant, ByRef DoSort As Boolean, ByRef SortColNum As Long, ByRef Ascending As Boolean, Recurse As Boolean)
1         On Error GoTo ErrHandler
          Dim ColumnTemplate2 As String
          Dim i As Long
          
2         If ColumnTemplate = vbNullString Then
3             If Recurse Then
4                 ColumnTemplate = "F"
5             Else
6                 ColumnTemplate = "N"
7             End If
8         End If
9         ColumnTemplate2 = Replace(ColumnTemplate, "^", vbNullString)
10        ColumnTemplate2 = Replace(UCase$(ColumnTemplate2), "V", vbNullString)
11        If ColumnTemplate2 <> ColumnTemplate Then
12            For i = 1 To Len(ColumnTemplate)
13                Select Case UCase$(Mid$(ColumnTemplate, i, 1))
                      Case "^"
14                        If i = 1 Then Throw "'^' (sort ascending) cannot be the first character in ColumnTemplate"
15                        If DoSort Then Throw "Sort characters ('^' or 'v') can appear only once in ColumnTemplate"
16                        DoSort = True
17                        Ascending = True
18                        SortColNum = i - 1
19                    Case "V"
20                        If i = 1 Then Throw "'v' (sort descending) cannot be the first character in ColumnTemplate"
21                        If DoSort Then Throw "Sort characters ('^' or 'v') can appear only once in ColumnTemplate"
22                        DoSort = True
23                        Ascending = False
24                        SortColNum = i - 1
25                End Select
26            Next i
27        End If

28        ReDim Headers(1 To 1, 1 To Len(ColumnTemplate2))

29        For i = 1 To Len(ColumnTemplate2)
30            Select Case UCase$(Mid$(ColumnTemplate2, i, 1))
                  Case "N"
31                    Headers(1, i) = "FileName"
32                Case "F"
33                    Headers(1, i) = "FullName"
34                Case "B"
35                    Headers(1, i) = "Attributes"
36                Case "D"
37                    Headers(1, i) = "Folder"
38                Case "R"
39                    Headers(1, i) = "RelativeName"
40                Case "S"
41                    Headers(1, i) = "Size"
42                Case "M"
43                    Headers(1, i) = "DateLastModified"
44                Case "#"
45                    Headers(1, i) = "MD5"
46                Case "C"
47                    Headers(1, i) = "DateCreated" 'Have to use FileSystemObject!
48                Case "A"
49                    Headers(1, i) = "DateLastAccessed" 'Have to use FileSystemObject!
50                Case "T"
51                    Headers(1, i) = "FileOrFolder"
52                Case Else
53                    Throw "Unrecognised character '" + Mid$(ColumnTemplate2, i, 1) + "'in ColumnTemplate. Allowed characters: N for Name, D for Folder, " & _
                          "F for FullName, R for RelativeName, S for Size, B for Attributes, # for MD5 Hash, M for DateLastModified, C for DateCreated, " & _
                          "A for DateLastAccessed, T for Type (File or folDer). Use '^' and 'v' after a character to sort."
54            End Select
55        Next i
          
56        ColumnTemplate = ColumnTemplate2
          
57        Exit Sub
ErrHandler:
58        Throw "#ParseColumnTemplate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegExpFromPattern
' Author     : Philip Swannell
' Date       : 15-May-2019
' Purpose    : Convert a Pattern as recognised by Dir (* for multiple characters and ? for single character) into a
'              regular expression. Previously was using the Like function to interpret the pattern but Like has more
'              powerful pattern matching than Dir
' Parameters :
'  Pattern:    The match pattern as Dir could interpret it
' -----------------------------------------------------------------------------------------------------------------------
Private Function RegExpFromPattern(Pattern As String)
          Dim i As Long
          Dim Res As String
          Const TheChars = ".$^{}[]()|+"
1         On Error GoTo ErrHandler

2         Res = Replace(Pattern, "\", "\\")
3         For i = 1 To Len(TheChars)
4             Res = Replace(Res, Mid$(TheChars, i, 1), "\" + Mid$(TheChars, i, 1))
5         Next i
6         Res = Replace(Res, "*", ".*")
7         Res = Replace(Res, "?", ".")
8         Res = "^" & Res & "$"

9         RegExpFromPattern = Res
10        Exit Function
ErrHandler:
11        Throw "#RegExpFromPattern (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
