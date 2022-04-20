Attribute VB_Name = "modSync"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SyncPGSOneDriveToSolumSharedDocuments
' Author     : Philip Swannell
' Date       : 16-Feb-2021
' Purpose    : Needed during lockdown so that SolumAddin etc can be easily available to other employees...
'              Call from button on Audit sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub SyncPGSOneDriveToSolumSharedDocuments()

          Const Target = "C:\Users\phili\Solum Financial Limited\Shared Documents - Documents\SolumSoftware\LatestVersion"

          Dim ReleaseFolder As String, TargetFolders
          Dim Res As VbMsgBoxResult
1         On Error GoTo ErrHandler
2         ReleaseFolder = DataFromAuditSheet("NetworkReleaseFolder")
3         If Right(ReleaseFolder, 1) = "\" Then ReleaseFolder = Left(ReleaseFolder, Len(ReleaseFolder) - 1)

          Dim Prompt As String
4         Prompt = "Folder to which new code is released:" & vbLf & _
              ReleaseFolder

5         Prompt = Prompt + vbLf + vbLf + _
              "Folder from which team members can install:" + vbLf + Target

6         Prompt = Prompt + vbLf + vbLf + "What would you like to do?"

7         Res = MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, , "Compare", "Synchronise Release -> Install", , , 500)

8         Select Case Res
              Case vbCancel
9                 Exit Sub
10            Case vbYes
11                g sFolderCompare(ReleaseFolder, Target, True)
12            Case vbNo
                  'Hand crafted code. SFolderSynchronise not quite able to do what we want.
                  'Filter out ISDA files and copy rather than sync
                  'BUT OneDrive commercial has a terrible behaviour that makes checking file hashes pointless!
                  'https://superuser.com/questions/1345933/why-is-onedrive-changing-the-size-of-files

                  Dim SourceFiles, TargetFiles, DirListRet, i As Long
13                DirListRet = ThrowIfError(sDirList(ReleaseFolder, True, False, "RF"))
                  'Const filter = "^((?!ISDA).)*$"
                  'ChooseVector = sIsRegMatch(filter, sSubArray(DirlistRet, 1, 1, , 1))
                  'DirlistRet = sMChoose(DirlistRet, ChooseVector)
14                SourceFiles = sSubArray(DirListRet, 1, 2, , 1)
15                TargetFiles = sJoinPath(Target, sSubArray(DirListRet, 1, 1, , 1))

16                TargetFolders = sSplitPath(TargetFiles, False)
17                TargetFolders = sRemoveDuplicates(TargetFolders)
18                sCreateFolder TargetFolders

19                ShowFileInSnakeTail

                  'SourceHashes = sFileCheckSum(SourceFiles)
                  'TargetHashes = sFileCheckSum(TargetFiles)
20                For i = 1 To sNRows(SourceFiles)
                      '   If SourceHashes(i, 1) <> TargetHashes(i, 1) Then
21                    MessageLogWrite "Copying to " & TargetFiles(i, 1)
22                    ThrowIfError sFileCopy(SourceFiles(i, 1), TargetFiles(i, 1))
                      '   End If
23                Next i
                  
24        End Select

25        Exit Sub
ErrHandler:
26        SomethingWentWrong "#SyncPGSOneDriveToSolumSharedDocuments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderSynchronise
' Author    : Philip Swannell
' Date      : 10-Oct-2018
' Purpose   : Mirror SourceFolder to Target folder, with minimal file copying: files in Source but not
'             Target are copied; files in Target but not Source are deleted; files common
'             to  Source and Target are copied if they differ.
' Arguments
' SourceFolder: The "source" folder. May or may not have terminating backslash.
' TargetFolder: The "target" folder. May or may not have terminating backslash. If this folder does not
'             yet exist it is created.
' FileFilter: Case-insensitive filter on file full names.
'             a) Pattern match, e.g.:  *.txt
'             b) 'RegExp' followed by a regular expression e.g.:  RegExp\.dll$|\.txt$
'             If a file exists in both Source and Target but does not match FileFilter then
'             it is deleted in Target.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderSynchronise(ByVal SourceFolder As String, ByVal TargetFolder As String, Optional FileFilter As String)
Attribute sFolderSynchronise.VB_Description = "Mirror SourceFolder to Target folder, with minimal file copying: files in Source but not Target are copied; files in Target but not Source are deleted; files common to  Source and Target are copied if they differ."
Attribute sFolderSynchronise.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim anyInTarget As Boolean
          Dim Common
          Dim FromFile As String
          Dim i As Long
          Dim inSNotT
          Dim InTNotS
          Dim SourceFiles
          Dim SourceShortNames
          Dim TargetFiles
          Dim TargetFolders
          Dim TargetShortNames
          Dim ThisFile As String
          Dim ToFile As String
          Dim NumDeletedInTarget As Long
          Dim NumIdentical As Long
          Dim NumNew As Long
          Dim NumChanged As Long
          Dim STK As clsStacker

1         On Error GoTo ErrHandler
2         Set STK = CreateStacker()

3         SourceFolder = Replace(SourceFolder, "/", "\")
4         TargetFolder = Replace(TargetFolder, "/", "\")
5         If Right$(SourceFolder, 1) <> "\" Then SourceFolder = SourceFolder + "\"
6         If Right$(TargetFolder, 1) <> "\" Then TargetFolder = TargetFolder + "\"

          'If TargetFolder is a sub-folder of SourceFolder then each call to the function would create deeper and deeper nested folders, so disallow, though test is not bomb-proof
7         If Len(TargetFolder) >= Len(SourceFolder) Then
8             If LCase$(Left$(TargetFolder, Len(SourceFolder))) = LCase$(SourceFolder) Then
9                 Throw "TargetFolder must not be a sub-folder of SourceFolder"
10            End If
11        End If

12        If Not sFolderExists(TargetFolder) Then
              Dim Res
13            Res = sCreateFolder(TargetFolder)
14            If sIsErrorString(Res) Then
15                Throw "TargetFolder ('" + TargetFolder + "') does not exist and cannot be created, with error " + Res
16            End If
17        End If

18        If Not sFolderExists(SourceFolder) Then
19            Throw "Cannot find Source Folder: '" & SourceFolder & "'"
20        Else
21            Debug.Print String(100, "-")
22            Debug.Print "Listing files in SourceFolder" ' '" + SourceFolder + "'"
23            SourceFiles = ThrowIfError(sDirList(SourceFolder, True, False, "FS", , FileFilter))
24            SourceShortNames = sArrayRight(sSubArray(SourceFiles, 1, 1, , 1), -Len(SourceFolder))
25            ThrowIfError sCreateFolder(TargetFolder)
26            Debug.Print "Listing files in TargetFolder" ' '" + TargetFolder + "'"
27            TargetFiles = ThrowIfError(sDirList(TargetFolder, True, True, "FS", , FileFilter))
28            If sNRows(TargetFiles) = 1 Then
29                anyInTarget = False
30            Else
31                anyInTarget = True
32                TargetFiles = sSubArray(TargetFiles, 2)
33            End If
34            TargetShortNames = sArrayRight(sSubArray(TargetFiles, 1, 1, , 1), -Len(TargetFolder))
35            TargetFolders = sRemoveDuplicates(sSplitPath(sArrayConcatenate(TargetFolder, SourceShortNames), False))
                
              Dim CreateFolderResult
36            CreateFolderResult = ThrowIfError(sFirstError(sCreateFolder(TargetFolders)))

37            If anyInTarget Then
38                InTNotS = sCompareTwoArrays(TargetShortNames, SourceShortNames, "12")
39                For i = 2 To sNRows(InTNotS)
40                    ThisFile = TargetFolder & InTNotS(i, 1)
41                    Debug.Print "Deleting " & ThisFile
42                    Res = sFileDelete(ThisFile)
43                    If sIsErrorString(Res) Then
44                        STK.Stack2D sArrayRange("Delete Failed", ThisFile, Res)
45                    Else
46                        NumDeletedInTarget = NumDeletedInTarget + 1
47                    End If
48                Next
49            End If

50            inSNotT = sCompareTwoArrays(TargetShortNames, SourceShortNames, "21")
51            For i = 2 To sNRows(inSNotT)
52                FromFile = SourceFolder & inSNotT(i, 1)
53                ToFile = TargetFolder & inSNotT(i, 1)
54                Debug.Print "Copying '" & FromFile & "' to '" & ToFile & "'"
55                Res = sFileCopy(FromFile, ToFile)
56                If sIsErrorString(Res) Then
57                    STK.Stack2D sArrayRange("Copy Failed", inSNotT(i, 1), Res)
58                Else
59                    NumNew = NumNew + 1
60                End If
61            Next

62            Common = sCompareTwoArrays(TargetShortNames, SourceShortNames, "Common")
63            For i = 2 To sNRows(Common)
64                FromFile = SourceFolder & Common(i, 1)
65                ToFile = TargetFolder & Common(i, 1)
66                If sFileInfo(FromFile, "Size") <> sFileInfo(ToFile, "Size") Then
67                    Debug.Print "Copying '" & FromFile & "' to '" & ToFile & "'"
68                    Res = sFileCopy(FromFile, ToFile)
69                    If sIsErrorString(Res) Then
70                        STK.Stack2D sArrayRange("Copy Failed", Common(i, 1), Res)
71                    Else
72                        NumChanged = NumChanged + 1
73                    End If
74                Else
                      Dim MD51 As String
                      Dim MD52 As String
                      '   MD5 calculation does not work on the IBM server used for the ISDA project, so be careful
75                    MD51 = sFileInfo(FromFile, "MD5")
76                    MD52 = sFileInfo(ToFile, "MD5")
77                    If MD51 <> MD52 Or sIsErrorString(MD51) Then
78                        Debug.Print "Copying '" & FromFile & "' to '" & ToFile & "'"
79                        Res = sFileCopy(FromFile, ToFile)
80                        If sIsErrorString(Res) Then
81                            STK.Stack2D sArrayRange("Copy Failed", Common(i, 1), Res)
82                        Else
83                            NumChanged = NumChanged + 1
84                        End If
85                    Else
86                        NumIdentical = NumIdentical + 1
87                    End If
88                End If
89            Next

90            Debug.Print "sFolderSynchronise. Source folder mirrored to target folder." + vbLf + _
                  "Source folder: '" + SourceFolder + "'" + vbLf + _
                  "Target folder: '" + TargetFolder + "'"
91            Debug.Print String(100, "-")
92        End If
          Dim Report
93        Report = sArrayRange(sArrayStack("SourceFolder", "TargetFolder", "New files copied", "Changed files copied", "Files deleted in Target", "Identical files not copied", "Failed operations"), _
              sArrayStack(SourceFolder, TargetFolder, NumNew, NumChanged, NumDeletedInTarget, NumIdentical, STK.NumRows))
94        If STK.NumRows > 0 Then
95            Report = sArrayStack(Report, STK.Report)
96        End If
              
97        sFolderSynchronise = Report
98        Exit Function
ErrHandler:
99        sFolderSynchronise = "#sFolderSynchronise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFolderCompare
' Author    : Philip Swannell
' Date      : 13-Nov-2020
' Purpose   : Compares two folders and lists files that are identical, files that differ and files that
'             appear in one folder but not the other.
' Arguments
' Folder1   : Full path to the first folder.
' Folder2   : Full path to the second folder.
' Recurse   : If TRUE then the function recurses into sub folders.
' -----------------------------------------------------------------------------------------------------------------------
Function sFolderCompare(ByVal Folder1 As String, ByVal Folder2 As String, Recurse As Boolean)
Attribute sFolderCompare.VB_Description = "Compares two folders and lists files that are identical, files that differ and files that appear in one folder but not the other."
Attribute sFolderCompare.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim List1 As Variant
          Dim List2 As Variant
          Dim FilesCommon As Variant
          Dim FilesIdentical As Variant
          Dim FilesDifferent As Variant
          Dim FilesIn1Not2 As Variant
          Dim FilesIn2Not1 As Variant
          Dim FilesUndetermined As Variant
          Dim Headers As Variant
          Dim NA As Variant
          Dim i As Long

1         On Error GoTo ErrHandler

          'If one but not the other is backslash-terminated make both backslash-terminated, necessary
          'thanks to behaviour of dirlist with "R" in the ColumnTemplate. Mmm is that behaviour desirable?..
2         Folder1 = Replace(Folder1, "/", "\")
3         Folder2 = Replace(Folder2, "/", "\")
4         If Right(Folder1, 1) = "\" And Right(Folder2, 1) <> "\" Then
5             Folder2 = Folder2 + "\"
6         ElseIf Right(Folder1, 1) <> "\" And Right(Folder2, 1) = "\" Then
7             Folder1 = Folder1 + "\"
8         End If

9         Headers = sArrayRange("Identical", "Different", "Not Determined(MD5 error)", "In1AndNotIn2", "In2AndNotIn1")
10        NA = CVErr(xlErrNA)
11        FilesIdentical = NA
12        FilesDifferent = NA
13        FilesUndetermined = NA
14        FilesIn1Not2 = NA
15        FilesIn2Not1 = NA
          'Note that it's inefficient to calculate the file hash of all files - would be better to only
          'calculate hashes for files that are common to both folders, but in practice that is likely to be most files
16        List1 = ThrowIfError(sDirList(Folder1, Recurse, True, "R#"))
17        List2 = ThrowIfError(sDirList(Folder2, Recurse, True, "R#"))

18        If sNRows(List1) = 1 And sNRows(List2) = 1 Then
19            sFolderCompare = Headers
20            Exit Function
21        ElseIf sNRows(List1) = 1 Then
22            FilesIn2Not1 = sSubArray(List2, 2, 1, , 1)
23        ElseIf sNRows(List2) = 1 Then
24            FilesIn1Not2 = sSubArray(List1, 2, 1, , 1)
25        Else
26            FilesIn1Not2 = sCompareTwoArrays(sSubArray(List1, 2, 1, , 1), sSubArray(List2, 2, 1, , 1), "12")
27            If sNRows(FilesIn1Not2) > 1 Then
28                FilesIn1Not2 = sSubArray(FilesIn1Not2, 2)
29            Else
30                FilesIn1Not2 = NA
31            End If
32            FilesIn2Not1 = sCompareTwoArrays(sSubArray(List1, 2, 1, , 1), sSubArray(List2, 2, 1, , 1), "21")
33            If sNRows(FilesIn2Not1) > 1 Then
34                FilesIn2Not1 = sSubArray(FilesIn2Not1, 2)
35            Else
36                FilesIn2Not1 = NA
37            End If
38            FilesCommon = sCompareTwoArrays(sSubArray(List1, 2, 1, , 1), sSubArray(List2, 2, 1, , 1), "Common")
39            If sNRows(FilesCommon) > 1 Then
40                FilesCommon = sSubArray(FilesCommon, 2)
                  Dim Hashes1, Hashes2, ChooserIdentical, ChooserDifferent, ChooserUndetermined
                  Dim AnyIdentical As Boolean, AnyDifferent As Boolean, AnyUndetermined As Boolean
41                Hashes1 = sVLookup(FilesCommon, List1)
42                Hashes2 = sVLookup(FilesCommon, List2)
43                ChooserIdentical = sReshape(False, sNRows(FilesCommon), 1)
44                ChooserDifferent = ChooserIdentical
45                ChooserUndetermined = ChooserIdentical
46                For i = 1 To sNRows(FilesCommon)
47                    If Left(Hashes1(i, 1), 1) = "#" Or Left(Hashes2(i, 1), 1) = "#" Then
48                        ChooserUndetermined(i, 1) = True
49                        AnyUndetermined = True
50                    ElseIf Hashes1(i, 1) = Hashes2(i, 1) Then
51                        ChooserIdentical(i, 1) = True
52                        AnyIdentical = True
53                    Else
54                        ChooserDifferent(i, 1) = True
55                        AnyDifferent = True
56                    End If
57                Next i
58                If AnyIdentical Then
59                    FilesIdentical = sMChoose(FilesCommon, ChooserIdentical)
60                End If
61                If AnyDifferent Then
62                    FilesDifferent = sMChoose(FilesCommon, ChooserDifferent)
63                End If
64                If AnyUndetermined Then
65                    FilesUndetermined = sMChoose(FilesCommon, ChooserUndetermined)
66                End If
67            End If
68        End If

69        sFolderCompare = sArrayStack(Headers, sArrayRange(FilesIdentical, FilesDifferent, FilesUndetermined, FilesIn1Not2, FilesIn2Not1))

70        Exit Function
ErrHandler:
71        sFolderCompare = "#sFolderCompare (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
