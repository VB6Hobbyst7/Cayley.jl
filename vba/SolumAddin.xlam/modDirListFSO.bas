Attribute VB_Name = "modDirListFSO"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modDirListOld
' Author    : Philip Swannell
' Date      : 14-June-2013
' Purpose   : Spreadsheet-callable function to return a list of all files in a directory.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module
Private m_DoFiles As Boolean
Private m_DoDirectories As Boolean
Private m_STK As clsStacker
Private m_ThisRow() As Variant
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDirListFSO
' Author    : Philip Swannell
' Date      : 14-Jun-2013
' Purpose   : Lists files or folders in a given folder(s). Can return either files or folders; can
'             recurse through the folder structure; can filter on file/folder names.
'             ColumnTemplate sets column contents and (optionally) the column on which the
'             return is sorted.
' Arguments
' Folder    : Path to folder e.g. "c:\temp" or equivalently "c:\temp\". May be an array to list files in
'             disparate folders - but see also Recurse argument.
' Recurse   : Logical value to determine whether the returned list includes files from sub folders or
'             not. When omitted defaults to FALSE, i.e. the return only shows the contents
'             of the folder, not of sub-folders.
' WithHeaders: Logical value to determine whether the returned array has a header row or not. If this
'             argument is omitted, then headers are not returned.
' ColumnTemplate: String to define columns. N = Name, F = FullName, S = Size, # = MD5 hash, M =
'             DateLastModified, A = DateLastAccessed, C = DateCreated, T = Type - (File or
'             Directory), v or ^ = return is sorted on prior column. If omitted, only the
'             file name is returned.
' FilesOrDirectories: String to determine if Files are listed (F) or Directories (D) or both (FD). If omitted,
'             then only files are listed.
' FileFilter: Optional case-insensitive filter on file names with path. Two syntaxes:
'             a) Pattern match (cf sArrayLike):    *.txt to match all text files
'             b) Regular Expression. 1st 6 characters must be RegExp.  RegExp\.dll$|\.txt$
'             to match either dll files or text files
' FolderFilter: Optional case-insensitive filter on folder names, with the same syntax as FileFilter.
'
' Example:    =sDirList("C:\Temp\",TRUE,TRUE,"FSv")
'             returns a two-column array, with a header row, of the full names and sizes of
'             all files in c:\Temp and is subdirectories. The return is sorted by file size
'             in descending order, i.e. largest file at the top.
' -----------------------------------------------------------------------------------------------------------------------
Public Function sDirListFSO(Folder As Variant, _
        Optional Recurse As Boolean, _
        Optional WithHeaders As Boolean = False, _
        Optional ByVal ColumnTemplate As String = "N", _
        Optional FilesOrDirectories As String = "F", _
        Optional ByVal FileFilter As String, _
        Optional ByVal FolderFilter As String)

1         On Error GoTo ErrHandler
          Dim Ascending As Boolean
          Dim DoSort As Boolean
          Dim i As Long
          Dim SortColNum As Long
          Dim Headers

2         If FunctionWizardActive() Then
3             sDirListFSO = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         ParseColumnTemplate ColumnTemplate, Headers, DoSort, SortColNum, Ascending, Recurse
        
7         ReDim m_ThisRow(1 To 1, 1 To Len(ColumnTemplate))

8         Set m_STK = CreateStacker()

9         If WithHeaders Then
10            m_STK.StackData Headers
11        End If

12        m_DoFiles = InStr(LCase$(FilesOrDirectories), "f") > 0
13        m_DoDirectories = InStr(LCase$(FilesOrDirectories), "d") > 0
14        If Not (m_DoFiles Or m_DoDirectories) Then
15            sDirListFSO = "#FilesOrDirectories must contain at least one of F (for Files) or D (for Directories)!"
16            Exit Function
17        End If

          Dim FileRegExp As VBScript_RegExp_55.RegExp
18        Set FileRegExp = New RegExp
19        If LCase$(Left$(FileFilter, 6)) = "regexp" Then
20            If VarType(sIsRegMatch(Mid$(FileFilter, 7), "Foo", False)) <> vbBoolean Then Throw "Regular Expression " + Mid$(FileFilter, 7) + " is not valid"
21            FileRegExp.Pattern = Mid$(FileFilter, 7)
22            FileRegExp.IgnoreCase = True
23            FileRegExp.Global = False
24            FileFilter = vbNullString
25        ElseIf VarType(SafeLike("Foo", FileFilter)) <> vbBoolean Then
26            Throw "FileFilter is not a valid string to match against"
27        End If

          Dim FolderRegExp As VBScript_RegExp_55.RegExp
28        Set FolderRegExp = New RegExp
29        If LCase$(Left$(FolderFilter, 6)) = "regexp" Then
30            If VarType(sIsRegMatch(Mid$(FolderFilter, 7), "Foo", False)) <> vbBoolean Then Throw "Regular Expression " + Mid$(FolderFilter, 7) + " is not valid"
31            FolderRegExp.Pattern = Mid$(FolderFilter, 7)
32            FolderRegExp.IgnoreCase = True
33            FolderRegExp.Global = False
34            FolderFilter = vbNullString
35        ElseIf VarType(SafeLike("Foo", FolderFilter)) <> vbBoolean Then
36            Throw "FolderFilter is not a valid string to match against"
37        End If

38        If VarType(Folder) = vbString Then
39            DirListCore CStr(Folder), Recurse, WithHeaders, ColumnTemplate, FilesOrDirectories, FileFilter, FolderFilter, FileRegExp, FolderRegExp
40        ElseIf IsArray(Folder) Then
41            If InStr(ColumnTemplate, "R") > 0 Then Throw "ColumnTemplate cannot contain 'R' (Relative Name) when Folder is given as an array of folders"
              Dim FirstCall As Boolean
              Dim thisFolder As Variant
42            FirstCall = True
43            For Each thisFolder In Folder
44                DirListCore CStr(thisFolder), Recurse, WithHeaders And FirstCall, ColumnTemplate, FilesOrDirectories, FileFilter, FolderFilter, FileRegExp, FolderRegExp
45                FirstCall = False
46            Next thisFolder
47        Else
48            Throw "Folder must be a string or an array of strings"
49        End If

          Dim Result As Variant
50        Result = m_STK.Report
51        If InStr(UCase$(ColumnTemplate), "R") > 0 Then
              Dim j As Long
52            For j = 1 To Len(ColumnTemplate)
53                If Mid$(ColumnTemplate, j, 1) = "R" Then
54                    For i = IIf(WithHeaders, 2, 1) To sNRows(Result)
55                        Result(i, j) = RelativeName(CStr(Result(i, j)), CStr(Folder))
56                    Next i
57                End If
58            Next j
59        End If

60        If DoSort Then
61            If WithHeaders Then
62                If sNRows(Result) > 1 Then
63                    Result = sArrayStack(sSubArray(Result, 1, 1, 1), sSortedArray(sSubArray(Result, 2), SortColNum, , , Ascending))
64                End If
65            Else
66                If Not sIsErrorString(Result) Then
67                    Result = sSortedArray(Result, SortColNum, , , Ascending)
68                End If
69            End If
70        End If

71        sDirListFSO = Result

72        Exit Function
ErrHandler:
73        sDirListFSO = "#sDirListFSO: line(" + CStr(Erl) + ") " + Err.Description + "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FilterMatches
' Author    : Philip Swannell
' Date      : 10-Sep-2013
' Purpose   : Encapsulate the testing of whether a file name matches a FilterString
'             could change to use Regular expressions or "*" and "?" syntax
'             21 Sep 2015 - changed by PGS to use Like operator - hence support * and ? wildcards.
'             9 Dec 2015 - changed by PGS to support Regular expressions. See help for sDirListFSO.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FilterMatches(FilterString As String, FilterRegExp As VBScript_RegExp_55.RegExp, StringToCheck As String) As Boolean
1         On Error GoTo ErrHandler

2         If FilterRegExp.Pattern <> vbNullString Then
3             FilterMatches = FilterRegExp.Test(StringToCheck)
4         ElseIf FilterString = vbNullString Then
5             FilterMatches = True
6         Else
7             FilterMatches = SafeLike(StringToCheck, FilterString)        'Use SafeLike since that ensures case-insensitive
8         End If

9         Exit Function
ErrHandler:
10        Throw "#FilterMatches (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DirListCore
' Author    : Philip Swannell
' Date      : 15-Jun-2013
' Purpose   : Subroutine of sDirListFSO
' -----------------------------------------------------------------------------------------------------------------------
Private Sub DirListCore(Folder As String, _
        Recurse As Boolean, _
        WithHeaders As Boolean, _
        ColumnTemplate As String, _
        FilesOrDirectories As String, _
        FileFilter As String, _
        FolderFilter As String, _
        FileRegExp As VBScript_RegExp_55.RegExp, _
        FolderRegExp As VBScript_RegExp_55.RegExp)

          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim myFile As Scripting.file
          Dim myFolder As Scripting.Folder
          Dim myFolder2 As Scripting.Folder
          Dim mySubFolder As Scripting.Folder
          Dim NumCols As Long
          Dim Offset As Long
          Static Hasher As Object
          Dim ZeroByteHash As String

1         If InStr(ColumnTemplate, "#") > 0 Then
2             If Hasher Is Nothing Then
3                 Set Hasher = CreateHasher(True, ZeroByteHash)
4             End If
5         End If

6         On Error GoTo ErrHandler

7         If WithHeaders Then Offset = 1

8         Set FSO = New Scripting.FileSystemObject
9         Set myFolder = FSO.GetFolder(Folder)
10        NumCols = Len(ColumnTemplate)

11        If m_DoDirectories Then
12            For Each myFolder2 In myFolder.SubFolders
13                If FilterMatches(FolderFilter, FolderRegExp, myFolder2.Name) Then
14                    For i = 1 To NumCols
15                        Select Case UCase$(Mid$(ColumnTemplate, i, 1))
                              Case "N"
16                                m_ThisRow(1, i) = myFolder2.Name
17                            Case "F", "R"
18                                m_ThisRow(1, i) = myFolder2.Path
19                            Case "S"
20                                m_ThisRow(1, i) = myFolder2.Size
21                            Case "M"
22                                m_ThisRow(1, i) = CDbl(myFolder2.DateLastModified)
23                            Case "A"
24                                m_ThisRow(1, i) = CDbl(myFolder2.DateLastAccessed)
25                            Case "C"
26                                m_ThisRow(1, i) = CDbl(myFolder2.DateCreated)
27                            Case "T"
28                                m_ThisRow(1, i) = "D"
29                            Case "#", "B", "D"
30                                m_ThisRow(1, i) = vbNullString
                              
31                        End Select
32                    Next i
33                    m_STK.StackData m_ThisRow
34                End If
35            Next myFolder2
36        End If

37        If m_DoFiles Then
38            For Each myFile In myFolder.Files
39                If FilterMatches(FileFilter, FileRegExp, myFile.Name) Then

40                    For i = 1 To NumCols
41                        Select Case UCase$(Mid$(ColumnTemplate, i, 1))
                              Case "N"
42                                m_ThisRow(1, i) = myFile.Name
43                            Case "F", "R"
44                                m_ThisRow(1, i) = myFile.Path
45                            Case "S"
46                                m_ThisRow(1, i) = myFile.Size
47                            Case "M"
48                                m_ThisRow(1, i) = CDbl(myFile.DateLastModified)
49                            Case "A"
50                                m_ThisRow(1, i) = CDbl(myFile.DateLastAccessed)
51                            Case "C"
52                                m_ThisRow(1, i) = CDbl(myFile.DateCreated)
53                            Case "T"
54                                m_ThisRow(1, i) = "F"
55                            Case "B"
56                                m_ThisRow(1, i) = myFile.Attributes
57                            Case "D"
58                                m_ThisRow(1, i) = myFolder.Path
59                            Case "#"
                                  Dim TheHash As String
60                                HashFromFileName Hasher, myFile.Path, TheHash, ZeroByteHash
61                                m_ThisRow(1, i) = TheHash
62                        End Select
63                    Next i
64                    m_STK.StackData m_ThisRow
65                End If
66            Next myFile
67        End If

68        If Recurse Then
69            For Each mySubFolder In myFolder.SubFolders
70                DirListCore mySubFolder.Path, Recurse, WithHeaders, ColumnTemplate, _
                      FilesOrDirectories, FileFilter, FolderFilter, FileRegExp, FolderRegExp
71            Next
72        End If

73        Exit Sub
ErrHandler:
74        Throw "#DirListCore: line(" + CStr(Erl) + ") " + Err.Description + "!"
End Sub

Private Function RelativeName(FullName As String, Folder As String)
1         If LCase$(Folder) = LCase$(Left$(FullName, Len(Folder))) Then
2             RelativeName = Mid$(FullName, Len(Folder) + 1)
3         Else
4             RelativeName = FullName
5         End If
End Function
