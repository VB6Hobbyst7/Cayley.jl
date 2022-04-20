Attribute VB_Name = "modRelease"
Option Explicit
Const gLocalWorkbooksFolder = "c:\SolumWorkbooks\"
Public Const gAddinName = "SolumAddin"    'As far as possible, soft code the use of the word "Solum"
Public Const gAddinName2 = "SolumSCRiPTUtils"    'As far as possible, soft code the use of the word "Solum"
Public Const gCompanyName = "Solum"    'As far as possible, soft code the use of the word "Solum"

Sub EmergencySaveMe()
1         ThisWorkbook.isAddin = True
2         sFileDelete "c:\temp\SolumAddin.xlam"
3         ThisWorkbook.SaveAs "c:\temp\SolumAddin.xlam", xlOpenXMLAddIn
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReleaseSoftware
' Author    : Philip Swannell
' Date      : 22-Jun-2016
' Purpose   : This method releases "all our software" i.e. the workbooks listed on the
'             static data worksheet of this addin plus our R source code and C sharp addin.
'             Release is to a "LatestVersion" folder on the network, together with a time-stamped backup.
' -----------------------------------------------------------------------------------------------------------------------
Sub ReleaseSoftware(Optional Comment As String)

1         On Error GoTo ErrHandler

          Dim Prompt As String
          Dim ReleaseToOneDrive As Variant
          Const Title = "Release " & gCompanyName & " Software"

          Const InstallerFileName = "Install.vbs"

          'Source folders...
2         Dim InstallerSourceFolder As String: InstallerSourceFolder = gBasePath & "Installer\"
3         Dim AddinsSourceFolder As String: AddinsSourceFolder = gBasePath & "Addins\"
          Dim RSourceSourceFolder As String
4         RSourceSourceFolder = Replace(gRSourcePath, "\", "/")

          Dim TestFolder1 As String
          Dim TestFolder1Good As Boolean
          Dim TestFolder2 As String
          Dim TestFolder2Good As Boolean

5         TestFolder1 = sRegistryRead("HKCU\SOFTWARE\Microsoft\OneDrive\UserFolder") + "\SolumSoftware\LatestVersion\"
6         TestFolder2 = "\\SOLUMSBS\Philip Shared\SolumSoftware\LatestVersion\"

7         If sFolderExists(TestFolder1) Then If sFolderIsWritable(TestFolder1) Then TestFolder1Good = True
8         If sFolderExists(TestFolder2) Then If sFolderIsWritable(TestFolder2) Then TestFolder2Good = True

9         If Not (TestFolder1Good Or TestFolder2Good) Then
10            Throw "Cannot release - no write access to either:" _
                  + vbLf + TestFolder1 + vbLf + "or" + vbLf + TestFolder2
11        ElseIf TestFolder1Good And Not TestFolder2Good Then
12            ReleaseToOneDrive = True
13        ElseIf (Not TestFolder1Good) And TestFolder2Good Then
14            ReleaseToOneDrive = False
15        Else
              Const chOneDrive = "Release to &OneDrive"
              Const chSSBS = "Release to &Network"
              Const chBoth = "Release to &Both OneDrive and Network"
              Dim Choice As String
16            Choice = ShowOptionButtonDialog(sArrayStack(chOneDrive, chSSBS, chBoth), , "Release to OneDrive or to SOLUMSBS?")
17            Select Case Choice
                  Case chOneDrive
18                    ReleaseToOneDrive = True
19                Case chSSBS
20                    ReleaseToOneDrive = False
21                Case chBoth
22                    ReleaseToOneDrive = sArrayStack(True, False)
23                Case Else
24                    Exit Sub
25            End Select
26        End If

27        Force2DArray ReleaseToOneDrive

          Dim AddinsTargetFolder As String
          Dim InstallerTargetFolder As String
          Dim LatestVersionFolderParent As String
          Dim NumReleases As Long
          Dim OldVersionsFolderParent As String
          Dim ReleaseCounter As Long
          Dim RSourceTargetFolder As String
          Dim WorkbooksTargetFolder As String
28        NumReleases = sNRows(ReleaseToOneDrive)

29        For ReleaseCounter = 1 To NumReleases

30            If ReleaseToOneDrive(ReleaseCounter, 1) Then
31                AddinsTargetFolder = TestFolder1 + "Addins\"
32                InstallerTargetFolder = TestFolder1 + "Installer\"
33                RSourceTargetFolder = TestFolder1 + "RSource\"
34                WorkbooksTargetFolder = TestFolder1 + "Workbooks\"
35                OldVersionsFolderParent = Replace(TestFolder1, "LatestVersion", "OldVersions")
36                LatestVersionFolderParent = TestFolder1
37            Else
38                AddinsTargetFolder = TestFolder2 + "Addins\"
39                InstallerTargetFolder = TestFolder2 + "Installer\"
40                RSourceTargetFolder = TestFolder2 + "RSource\"
41                WorkbooksTargetFolder = TestFolder2 + "Workbooks\"
42                OldVersionsFolderParent = Replace(TestFolder2, "LatestVersion", "OldVersions")
43                LatestVersionFolderParent = TestFolder2
44            End If

              Dim AllTargetFolders
              Dim ChooseVector As Variant
              Dim ChooseVectorForCopying As Variant
              Dim CopyFromList As Variant
              Dim CopyToList As Variant
              Dim i As Long
              Dim OldVersionsFolder As String
              Dim TheseFileNames As Variant
              Dim TimeStamp As String

45            If ThisWorkbook.VBProject.Protection = 1 Then
46                MsgBoxPlus "Please unlock the VBA code of " & gAddinName
47                Application.SendKeys "%{F11}%W{UP}{RETURN}"
48                Exit Sub
49            End If

50            TimeStamp = Format$(Now(), "yyyy-mm-dd hh-mm")

51            OldVersionsFolder = OldVersionsFolderParent + TimeStamp + Comment + "\"

52            TheseFileNames = sExpandDown(RangeFromSheet(shWorkbookLists, "AddinList"))
53            CopyFromList = sArrayConcatenate(AddinsSourceFolder, TheseFileNames)
54            CopyToList = sArrayConcatenate(AddinsTargetFolder, TheseFileNames)

55            TheseFileNames = InstallerFileName
56            CopyFromList = sArrayStack(CopyFromList, sArrayConcatenate(InstallerSourceFolder, TheseFileNames))
57            CopyToList = sArrayStack(CopyToList, sArrayConcatenate(InstallerTargetFolder, TheseFileNames))

58            TheseFileNames = sDirList(RSourceSourceFolder, False, False, "N")
59            CopyFromList = sArrayStack(CopyFromList, sArrayConcatenate(RSourceSourceFolder, TheseFileNames))
60            CopyToList = sArrayStack(CopyToList, sArrayConcatenate(RSourceTargetFolder, TheseFileNames))

61            TheseFileNames = sExpandDown(RangeFromSheet(shWorkbookLists, "WorkbookList"))
62            CopyFromList = sArrayStack(CopyFromList, sArrayConcatenate(gLocalWorkbooksFolder, TheseFileNames))
63            CopyToList = sArrayStack(CopyToList, sArrayConcatenate(WorkbooksTargetFolder, TheseFileNames))

              'Check SourceFiles exist
64            ChooseVector = sReshape(False, sNRows(CopyFromList), 1)
65            For i = 1 To sNRows(CopyFromList)
66                If Not sFileExists(CopyFromList(i, 1)) Then ChooseVector(i, 1) = True
67            Next i

              Dim NumNotFound
68            NumNotFound = sArrayCount(ChooseVector)
69            If NumNotFound > 0 Then
70                Prompt = "Cannot find the following " + IIf(NumNotFound = 1, "file", CStr(NumNotFound) + " files") + _
                      ". Do you want to proceed with the release anyway?" + vbLf + vbLf + _
                      sConcatenateStrings(sMChoose(CopyFromList, ChooseVector), vbLf)

71                If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel + vbDefaultButton2, Title, "Yes, Release anyway", "No, abort release", , , 500) <> vbOK Then Exit Sub
72            End If

              'Check target folders exist
73            AllTargetFolders = sArrayStack(AddinsTargetFolder, InstallerTargetFolder, RSourceTargetFolder, WorkbooksTargetFolder)
74            ChooseVector = sReshape(False, sNRows(AllTargetFolders), 1)
75            For i = 1 To sNRows(AllTargetFolders)
76                If Not sFolderExists(AllTargetFolders(i, 1)) Then ChooseVector(i, 1) = True
77            Next i

78            Select Case sArrayCount(ChooseVector)
                  Case 0
79                Case 1: Throw "Cannot find folder " + sMChoose(AllTargetFolders, ChooseVector)(1, 1), True
80                Case Else: Throw "Cannot find the following folders:" + vbLf + sConcatenateStrings(sMChoose(AllTargetFolders, ChooseVector), vbLf), True
81            End Select

              'Check target folders are writable
82            ChooseVector = sReshape(False, sNRows(AllTargetFolders), 1)
83            For i = 1 To sNRows(AllTargetFolders)
84                If Not sFolderIsWritable(AllTargetFolders(i, 1)) Then ChooseVector(i, 1) = True
85            Next i
86            Select Case sArrayCount(ChooseVector)
                  Case 0
87                Case 1: Throw "Cannot write to folder " + sMChoose(AllTargetFolders, ChooseVector)(1, 1), True
88                Case Else: Throw "Cannot write to the following folders:" + vbLf + sConcatenateStrings(sMChoose(AllTargetFolders, ChooseVector), vbLf), True
89            End Select

              'Check TargetFile (if it exists) is not newer than SourceFile, but don't worry if its newer but identical
90            ChooseVector = sReshape(False, sNRows(CopyFromList), 1)
              Dim FileList
              Dim Headers
91            Headers = sArrayRange("CopyFrom", "LastModifiedDate", "CopyTo", "LastModifiedDate", "TargetIsNewerAndDifferent")
92            FileList = sArrayRange(CopyFromList, sFileLastModifiedDate(CopyFromList), CopyToList, sFileLastModifiedDate(CopyToList))

93            For i = 1 To sNRows(CopyFromList)
94                If IsNumberOrDate(FileList(i, 4)) Then
95                    If FileList(i, 4) > FileList(i, 2) Then
96                        If sFileCheckSum(FileList(i, 3)) <> sFileCheckSum(FileList(i, 1)) Then
97                            ChooseVector(i, 1) = True
98                        End If
99                    End If
100               End If
101           Next i

102           If sArrayCount(ChooseVector) >= 1 Then
                  Dim DataToPaste As Variant
                  Dim Res
103               DataToPaste = sArrayStack(Headers, sArrayRange(FileList, ChooseVector))
104               g DataToPaste, ExMthdSpreadsheet

105               Prompt = "Warning:" + vbLf + "Files exist in the 'LatestVersions' folder on the network AND ARE NEWER than the version to be copied from this PC. Details have been pasted to a new workbook." + vbLf + vbLf + "What do you want to do?"
                  Const ChAbort = "Abort the release so that I can examine the list of problems"
                  Const chCautious = "Do the release but do not overwrite newer files"
                  Const chBold = "Do the release and overwrite newer files"
                  Const chHandCrafted = "Let me choose file-by-file"

106               Res = ShowOptionButtonDialog(sArrayStack(ChAbort, chCautious, chBold, chHandCrafted), "Release " & gCompanyName & " Software", Prompt, ChAbort)
107               If IsEmpty(Res) Then Exit Sub
108               If Res = ChAbort Then
109                   Exit Sub
110               ElseIf Res = chBold Then
111                   ChooseVectorForCopying = sReshape(True, sNRows(CopyFromList), 1)
112               ElseIf Res = chCautious Then
113                   ChooseVectorForCopying = sArrayNot(ChooseVector)
114               ElseIf Res = chHandCrafted Then
                      Dim FileList2
115                   FileList2 = sArrayRange(CopyFromList, sSubArray(FileList, 1, 2, , 1), sSubArray(FileList, 1, 4, , 1))
116                   FileList2 = sMChoose(FileList2, ChooseVector)
117                   For i = 1 To sNRows(FileList2)
118                       FileList2(i, 2) = Format$(FileList2(i, 2), "d-mmm-yyyy hh:mm:ss")
119                       FileList2(i, 3) = Format$(FileList2(i, 3), "d-mmm-yyyy hh:mm:ss")
120                   Next i
                      Dim TheOptions
121                   TheOptions = sJustifyArrayOfStrings(FileList2, "Tahoma", 8, vbTab)
122                   Prompt = "Files below are newer in the network release area." + vbLf + vbLf + "Please select which to copy."

123                   Res = ShowMultipleChoiceDialog(TheOptions, , , Prompt)
124                   If sArraysIdentical(Res, "#User Cancel!") Then Exit Sub
125                   If IsEmpty(Res) Then
126                       ChooseVectorForCopying = sArrayNot(ChooseVector)
127                   Else

                          Dim MatchIDs
                          Dim NewerFilesToCopy
128                       MatchIDs = sMatch(Res, TheOptions)
129                       Force2DArray MatchIDs
130                       NewerFilesToCopy = CreateMissing()
131                       For i = 1 To sNRows(MatchIDs)
132                           NewerFilesToCopy = sArrayStack(NewerFilesToCopy, FileList2(MatchIDs(i, 1), 1))
133                       Next i
134                       ChooseVectorForCopying = sArrayOr(sArrayNot(ChooseVector), sArrayIsNumber(sMatch(CopyFromList, NewerFilesToCopy)))
135                   End If
136               End If
137           Else
138               ChooseVectorForCopying = sReshape(True, sNRows(CopyFromList), 1)
139           End If

              Dim CopyResults
              Dim SomeFailed As Boolean
140           CopyResults = sReshape(vbNullString, sNRows(CopyFromList), 1)

              'Do the copy to LatestVersion
141           For i = 1 To sNRows(CopyFromList)
142               If ChooseVectorForCopying(i, 1) Then
143                   sCreateFolder sSplitPath(CopyToList(i, 1), False)
144                   CopyResults(i, 1) = sFileCopy(CopyFromList(i, 1), CopyToList(i, 1))
145                   If VarType(CopyResults(i, 1)) = vbString Then
146                       SomeFailed = True
147                   End If
148               Else
149                   CopyResults(i, 1) = "User decided not to copy"
150               End If
151           Next i

152           If SomeFailed Then
153               g sArrayStack(sArrayRange("SourceFile", "TargetFile", "CopyResult"), sArrayRange(CopyFromList, CopyToList, CopyResults)), ExMthdSpreadsheet
154               MsgBoxPlus "Some file copy commands failed. The active workbook lists the failures." + vbLf + vbLf + _
                      "Copy of latest version folder to old version folder will not be done", vbCritical
155           Else
156               ThrowIfError sFolderCopy(LatestVersionFolderParent, OldVersionsFolder)
157               Prompt = "Release " + CStr(ReleaseCounter) + " of " + CStr(sNRows(ReleaseToOneDrive)) + " has finished. " + CStr(sArrayCount(ChooseVectorForCopying)) + " files released"
158               If sArrayCount(sArrayNot(ChooseVectorForCopying)) > 0 Then
159                   Prompt = Prompt + ", " + CStr(sArrayCount(sArrayNot(ChooseVectorForCopying))) + " files not released."
160               Else
161                   Prompt = Prompt + "."
162               End If
163               MsgBoxPlus Prompt, vbInformation, Title, , , , , , , , IIf(ReleaseCounter < NumReleases, 30, 0), vbOK
164           End If

165       Next ReleaseCounter

166       Exit Sub
ErrHandler:
167       SomethingWentWrong "#ReleaseSoftware (line " & CStr(Erl) & "): " & Err.Description & "!", , Title
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetComment
' Author    : Philip Swannell
' Date      : 13-May-2016
' Purpose   : When we back-up files either the RSource of a "working snapshot" it can be
'             helpful to include some text in the directory name, such as "For testing only", "CapFloor2 not yet working" etc
' -----------------------------------------------------------------------------------------------------------------------
Sub GetComment(ByRef TheComment As String)
          Static Comment As String        'make static so that at least in one session we remember what we wrote last time...
1         On Error GoTo ErrHandler
2         Comment = InputBoxPlus("Short release comment (to be included in directory name)", "Release " & gCompanyName & " Software", Comment, , , 300)
3         Comment = Replace(Comment, vbCr, " ")
4         Comment = Replace(Comment, vbLf, " ")
5         Comment = sRegExReplace(Comment, "\\|/|:|\*|\?|""|<|>|\|", "_")        'replace characters illegal in file names
6         If Comment <> "False" Then If Len(Comment) > 1 Then If Left$(Comment, 1) <> " " Then Comment = " " + Comment
7         Do While Right$(Comment, 1) = " "        'don't want trailing spaces it screws up a subsequent call to sFolderCopy, didn't investigate why
8             Comment = Left$(Comment, Len(Comment) - 1)
9         Loop
10        AddReleaseCommentToMRU Comment
11        TheComment = Comment
12        Exit Sub
ErrHandler:
13        Throw "#GetComment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReleaseToNAS
' Author    : Philip
' Date      : 24-Dec-2015
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub ReleaseToNAS(wb As Excel.Workbook, SilentMode As Boolean)
1         On Error GoTo ErrHandler
          Dim TargetFolder1 As String
          Dim TargetFolder2 As String
          Dim Prompt As String

2         If Not SilentMode Then
3             Prompt = "Release " + wb.Name + "?" + vbLf + vbLf + "ReleaseComment is:" + vbLf + String(100, "-") + vbLf + ReleaseCommentFromAuditSheet(wb.Worksheets("Audit")) + vbLf + String(100, "-")
4             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, "Release " + wb.Name, "Release", , , , 400) <> vbOK Then Exit Sub
5         End If

6         TargetFolder1 = DataFromAuditSheet("NetworkReleaseFolder")
7         If Not sFolderIsWritable(TargetFolder1) Then Throw "Cannot release because folder '" + TargetFolder1 + " is not writeable"
8         TargetFolder2 = TargetFolder1 & "Addins\"
9         If Not sFolderIsWritable(TargetFolder2) Then Throw "Cannot release because folder '" + TargetFolder2 + " is not writeable"
10        CheckForOverwrite wb, True, False
11        SaveAddinAndExportVBA wb
12        AddReleaseCommentFromAuditSheetToMRU wb.Worksheets("Audit")
13        On Error GoTo ErrHandler
14        ThrowIfError sFileCopy(wb.FullName, TargetFolder2 + wb.Name)
15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#ReleaseToNAS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExportModules
' Author    : Philip Swannell
' Date      : 18-Dec-2015
' Purpose   : Exports the VBA code of a workbook to a folder
' -----------------------------------------------------------------------------------------------------------------------
Sub ExportModules(wb As Excel.Workbook, ByVal Folder As String, SaveWorkbookAlso As Boolean)
          Dim bExport As Boolean
          Dim c As VBIDE.VBComponent
          Dim FileName As String
          Dim STK As clsStacker
          Dim VersionNumber As Long

1         On Error GoTo ErrHandler
            
2         If wb.VBProject.Protection = 1 Then
3             Throw "VBProject is protected"
4             Exit Sub
5         End If

6         If Folder = vbNullString Then Throw "Folder must be provided"

7         If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"

8         If Not sFolderExists(Folder) Then Throw "Cannot find folder: '" + Folder + "'"

9         On Error Resume Next
10        Kill Folder & "*.bas*"
11        Kill Folder & "*.cls*"
12        Kill Folder & "*.frm*"
13        Kill Folder & "*.frx*"
14        On Error GoTo ErrHandler

15        Set STK = CreateStacker()
16        For Each c In wb.VBProject.VBComponents
17            bExport = True
18            FileName = c.Name

19            Select Case c.Type
                  Case vbext_ct_ClassModule
20                    FileName = FileName & ".cls"
21                Case vbext_ct_MSForm
22                    FileName = FileName & ".frm"
23                Case vbext_ct_StdModule
24                    FileName = FileName & ".bas"
25                Case vbext_ct_Document
26                    If c.CodeModule.CountOfLines <= 2 Then        'Only export sheet module if it contains code. Test CountOfLines <= 2 likely to be good enough in practice -
27                        bExport = False
28                    Else
29                        bExport = True
30                        FileName = FileName & ".cls"
31                    End If
32                Case Else
33                    bExport = False
34            End Select

35            If bExport Then
36                c.Export Folder & FileName
37                STK.Stack0D FileName
38            End If
39        Next c

40        STK.Stack0D wb.Name

41        On Error Resume Next
42        Kill Folder & "*.frx"        'These are binary files that we don't want to check in to Git
43        On Error GoTo ErrHandler

44        If IsInCollection(wb.Worksheets, "Audit") Then
45            If IsInCollection(wb.Worksheets("Audit").Names, "Headers") Then
                  Dim AuditSheetComments, j As Long, NR As Long, NC As Long
46                AuditSheetComments = sExpandDown(wb.Worksheets("Audit").Range("Headers")).Value
47                VersionNumber = AuditSheetComments(2, 1)
48                NR = sNRows(AuditSheetComments)
49                NC = sNCols(AuditSheetComments)
                  'Flip the contents of the Time column to be readable in text Format
                  Dim i As Long
                  Dim TimeCol As Variant
50                TimeCol = sMatch("Time", sArrayTranspose(sSubArray(AuditSheetComments, 1, 1, 1)))
51                If IsNumber(TimeCol) Then
52                    For i = 2 To NR
53                        If IsNumber(AuditSheetComments(i, TimeCol)) Then
54                            If AuditSheetComments(i, TimeCol) <= 1 Then
55                                If AuditSheetComments(i, TimeCol) >= 0 Then
56                                    AuditSheetComments(i, TimeCol) = Format$(AuditSheetComments(i, TimeCol), "hh:mm")
57                                End If
58                            End If
59                        End If
60                    Next i
61                End If

                  'Git treats unicode files as binary, so we can't save as unicode, so we have to replace high ascii by ?
                  Dim t1 As Double
62                t1 = sElapsedTime
63                For i = 1 To NR
64                    For j = 1 To NC
65                        If VarType(AuditSheetComments(i, j)) = vbString Then
66                            RemoveHighAscii AuditSheetComments(i, j)
67                        End If
68                    Next
69                Next
70                ThrowIfError sFileSave(Folder & "AuditSheetComments.txt", AuditSheetComments, vbTab, " ", True)
71                STK.Stack0D "AuditSheetComments.txt"
72            End If
73        End If

74        If SaveWorkbookAlso Then
              Dim TargetName As String, DotAt As Long
75            If VersionNumber = 0 Then
76                TargetName = wb.Name
77            Else
78                DotAt = InStrRev(wb.Name, ".")
79                TargetName = Left(wb.Name, DotAt - 1) + "_v" & CStr(VersionNumber) & Mid(wb.Name, DotAt)
80            End If

81            ThrowIfError sFileCopy(wb.FullName, Folder + TargetName)
82        End If
          'Save a listing of the files in the project
83        ThrowIfError sFileSave(Folder & "FilesInProject.txt", sSortedArray(STK.Report))

84        Exit Sub
ErrHandler:
85        Throw "#ExportModules (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function RemoveHighAscii(ByRef Str)
          Dim i As Long
1         For i = 1 To Len(Str)
2             If AscW(Mid(Str, i, 1)) > 255 Then
3                 Mid(Str, i, 1) = "?"
4             End If
5         Next
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReleaseWorkbookNewStyle
' Author     : Philip Swannell
' Date       : 19-Oct-2021
' Purpose    : New style of release code for workbooks not intended to be installed with "SolumWorkbooks" but instead
'              are in their own GitHub repo, example: VBAInterop.xlsm and VBA-CSV.xlsm
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub ReleaseWorkbookNewStyle(wb As Excel.Workbook)
          Dim Path As String
          Dim ExportVBATo As String
          Dim ParentDirectory As String
          Dim ChildDirectory As String
          Dim GrandchildDirectory As String
          Dim NowString As String
          Dim Prompt As String
          Dim Caption2 As String
          Const Title = "Release Workbook"
          
1         On Error GoTo ErrHandler

2         ParentDirectory = DataFromAuditSheet("ParentForBackupOfOtherWorkbooks")
3         ChildDirectory = ParentDirectory + sStringBetweenStrings(wb.Name, , ".") + "Backups\"
4         Path = wb.Path
5         If LCase(Left(Path, 11)) <> "c:\projects" Then
6             Throw "Workbook must be in a sub-folder of C:\Projects"
7         End If

8         ExportVBATo = WorkbookToLocalGitFolder(wb)

9         Prompt = "Release '" + wb.Name + "'?" + vbLf + vbLf + "Release does the following:" + vbLf + _
              "1) Saves the workbook to its current location" + vbLf + _
              "    (" + wb.Path + ")" + vbLf + _
              "2) Copies the workbook and its VBA to a time-stamped archive " + vbLf + _
              "    (" + ChildDirectory + "yyyy-mm-dd hh-mm-ss\)" + vbLf + _
              "3) Exports the workbook's VBA to the source-control directory" + vbLf + _
              "    (" + ExportVBATo + ")"

10        If wb.VBProject.Protection = 1 Then
11            Prompt = Prompt + vbLf + vbLf + "But first you must unlock the VBA code in the VBEditor (Alt F11)"
12            Caption2 = "Unlock the VBA code"
13        Else
14            Caption2 = "Yes, release"
15        End If

16        If MsgBoxPlus(Prompt, vbYesNo + vbQuestion, Title, Caption2, "No, quit", , , 700) <> vbYes Then Exit Sub
17        If wb.VBProject.Protection = 1 Then
              'Activate the VBIDE and close windows
18            Application.SendKeys "%{F11}%W{UP}{RETURN}"
19            Exit Sub
20        End If

21        If Not sFolderExists(ExportVBATo) Then
22            Throw "Cannot find folder for VBA export: '" + ExportVBATo + "'"
23        End If
24        If wb.VBProject.Protection = 1 Then
              'Activate the VBIDE and close windows
25            Application.SendKeys "%{F11}%W{UP}{RETURN}"
26            Exit Sub
27        End If

28        NowString = Format$(Now(), "yyyy-mm-dd hh-mm-ss")
29        GrandchildDirectory = ChildDirectory + NowString + "\"
30        ThrowIfError sCreateFolder(GrandchildDirectory)

          Dim isAddin As Boolean
31        isAddin = wb.isAddin
32        If isAddin = False Then
33            If LCase(Right(wb.Name, 5)) = ".xlam" Then
34                wb.isAddin = True
35            End If
36        End If
37        wb.Save
38        If wb.isAddin <> isAddin Then wb.isAddin = isAddin
39        ExportModules wb, GrandchildDirectory, True
40        ExportModules wb, ExportVBATo, False ' Save then need to use GitHub desktop to commit to Git, too painful to try automating

          'PGS 24 Jan 2022 Not sure if this is a good idea or not, but also copy to the location that old-style release released to
          Dim LatestVersionDirectory
41        LatestVersionDirectory = WorkbookReleaseFolder(wb.Name)
42        ThrowIfError sFileCopy(wb.FullNameURLEncoded, LatestVersionDirectory + wb.Name)

43        Exit Sub
ErrHandler:
44        Throw "#ReleaseWorkbookNewStyle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReleaseWorkbook
' Author    : Philip Swannell
' Date      : 01-Mar-2016
' Purpose   : General-purpose routine for releasing a workbook.
'        0)   Assumes any "release cleanup" code has already been run.
'        1)   Saves the workbook and VBA code to a time-stamped directory.
'             If ParentDirectory is \\Foo\Bar\ and workbook is called MyBook.xlsm then the
'             book and code files are saved to \\Foo\Bar\MyBookBackups\yyyy-mm-dd hh-mm-ss\
'        2)   Exports the VBA code for checking in to Git
'        3)   Saves workbook to "LatestVersions" directory
' -----------------------------------------------------------------------------------------------------------------------
Sub ReleaseWorkbook(wb As Excel.Workbook, Optional ByVal ParentDirectory As String, Optional TimeNow As Double, Optional SilentMode As Boolean)
          Const Title = "Release Workbook"
          Dim Caption2 As String
          Dim ChildDirectory As String
          Dim GrandchildDirectory As String
          Dim LatestVersionDirectory As String
          Dim NowString As String
          Dim Prompt As String

1         On Error GoTo ErrHandler

2         If LCase(Left(wb.FullName, 11)) = "c:\projects" Then
3             ReleaseWorkbookNewStyle wb
4             Exit Sub
5         End If

6         If LCase$(Left$(wb.FullName, Len(gLocalWorkbooksFolder))) <> LCase$(gLocalWorkbooksFolder) Then
7             Throw "Workbook must be saved to " + gLocalWorkbooksFolder + " (or a subfolder) before being released"
8         End If

9         If Not SilentMode Then
10            CheckForOverwrite wb, False
11        End If

12        If ParentDirectory = vbNullString Then ParentDirectory = DataFromAuditSheet("ParentForBackupOfOtherWorkbooks")
13        LatestVersionDirectory = WorkbookReleaseFolder(wb.Name)
14        ChildDirectory = ParentDirectory + sStringBetweenStrings(wb.Name, , ".") + "Backups\"

15        Prompt = "Release '" + wb.Name + "'?" + vbLf + vbLf + "Release does the following:" + vbLf + _
              "1) Saves the workbook to its current location" + vbLf + _
              "    (" + wb.Path + ")" + vbLf + _
              "2) Copies it to the LatestVersion directory" + vbLf + _
              "    (" + LatestVersionDirectory + ")" + vbLf + _
              "3) Copies the workbook and its VBA to a time-stamped archive " + vbLf + _
              "    (" + ChildDirectory + "yyyy-mm-dd hh-mm-ss\)" + vbLf + _
              "4) Exports the workbook's VBA to the source-control directory" + vbLf + _
              "    (" + WorkbookToLocalGitFolder(wb) + ")"

16        If wb.VBProject.Protection = 1 Then
17            Prompt = Prompt + vbLf + vbLf + "But first you must unlock the VBA code in the VBEditor (Alt F11)"
18            Caption2 = "Unlock the VBA code"
19        Else
20            Caption2 = "Yes, release"
21        End If

22        If Not SilentMode Then If MsgBoxPlus(Prompt, vbYesNo + vbQuestion, Title, Caption2, "No, quit", , , 600) <> vbYes Then Exit Sub
23        If wb.VBProject.Protection = 1 Then
              'Activate the VBIDE and close windows
24            Application.SendKeys "%{F11}%W{UP}{RETURN}"
25            Exit Sub
26        End If

27        If Not sFolderExists(LatestVersionDirectory) Then Throw "Cannot access " + LatestVersionDirectory, True
28        If Not sFolderIsWritable(LatestVersionDirectory) Then Throw "You do not have write access to " + LatestVersionDirectory, True

29        If Not sFolderExists(ParentDirectory) Then Throw "Cannot find folder " + ParentDirectory, True
30        If Not sFolderExists(ChildDirectory) Then
31            Prompt = "Directory " + ChildDirectory + " to which time-stamped versions of this workbook and its code modules will be saved does not exist. Would you like to create it now?"
32            If Not SilentMode Then If MsgBoxPlus(Prompt, vbYesNo + vbQuestion, Title, "Yes, proceed", "No, just quit") <> vbYes Then Exit Sub
33            ThrowIfError sCreateFolder(ChildDirectory)
34        End If
35        If TimeNow = 0 Then
36            NowString = Format$(Now(), "yyyy-mm-dd hh-mm-ss")
37        Else
38            NowString = Format$(TimeNow, "yyyy-mm-dd hh-mm-ss")
39        End If
40        GrandchildDirectory = ChildDirectory + NowString + "\"
41        ThrowIfError sCreateFolder(GrandchildDirectory)

42        wb.Save
43        ThrowIfError sFileCopy(wb.FullNameURLEncoded, GrandchildDirectory + wb.Name)
44        If LatestVersionDirectory <> ParentDirectory Then
45            ThrowIfError sFileCopy(wb.FullNameURLEncoded, LatestVersionDirectory + wb.Name)
46        End If

47        ExportModules wb, WorkbookToLocalGitFolder(wb), False      'For Git
48        ExportModules wb, GrandchildDirectory, True

49        If IsInCollection(wb.Worksheets, "Audit") Then
50            AddReleaseCommentFromAuditSheetToMRU wb.Worksheets("Audit")
51        End If

          'Update list of released workbooks
52        If Not SilentMode Then
              Dim RelativeName As String
53            If LCase$(Left$(wb.FullName, Len(gLocalWorkbooksFolder))) <> LCase$(gLocalWorkbooksFolder) Then Throw "Assertion failed, workbook should be in '" + gLocalWorkbooksFolder + "' (or a sub folder)"
54            RelativeName = Mid$(wb.FullName, Len(gLocalWorkbooksFolder) + 1)
              Dim R As Range
              Dim ReleaseComment As Variant
55            Set R = sExpandDown(shWorkbookLists.Range("WorkbookList"))
56            If Not IsNumber(sMatch(RelativeName, R.Value)) Then
                  Dim SPH As clsSheetProtectionHandler
57                Set SPH = CreateSheetProtectionHandler(shWorkbookLists)
58                R.Cells(R.Rows.Count + 1).Value = RelativeName
59                Prompt = "This seems to be the first time that the workbook '" + RelativeName + "' has been released." + vbLf + vbLf + _
                      gAddinName & " contains a list of all released workbooks that's used to create backups of all released workbooks. The list of workbooks has been updated and therefore " & gAddinName & " itself needs to be released." + vbLf + vbLf + _
                      "Would you like to release " & gAddinName & " now?"
60                If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, "Release " & gAddinName & "?") = vbOK Then
61                    ReleaseComment = "Added workbook '" + RelativeName + "' to Workbooklist on sheet WorkbookLists."
62                    ReleaseComment = InputBoxPlus("Please enter release comment:", "Release " & gAddinName & vbNullString, CStr(ReleaseComment), , , 400, 40)
63                    If VarType(ReleaseComment) = vbBoolean Then Exit Sub
64                    AddLineToAuditSheet shAudit
65                    RangeFromSheet(shAudit, "Headers").Cells(2, 5).Value = ReleaseComment
66                    ReleaseToNAS ThisWorkbook, True
67                    ThisWorkbook.isAddin = True
68                End If
69            End If
70        End If

71        Exit Sub
ErrHandler:
72        SomethingWentWrong "#ReleaseWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WorkbookReleaseFolder
' Author     : Philip Swannell
' Date       : 11-Dec-2017
' Purpose    : Encapsulate location of release folder for workbooks, including special handling for the many workbooks that
'              are involved in our work for ISDA
' Parameters :
'  WorkbookName:
' -----------------------------------------------------------------------------------------------------------------------
Function WorkbookReleaseFolder(WorkbookName As String)
          Dim BaseFolder As String
          Dim i As Long
1         On Error GoTo ErrHandler

2         BaseFolder = DataFromAuditSheet("NetworkReleaseFolder") + "Workbooks\"
3         If InStr(LCase$(WorkbookName), "isda simm") > 0 Then
4             For i = 2017 To 2037
5                 If InStr(LCase$(WorkbookName), "isda simm " + CStr(i)) > 0 Then
6                     WorkbookReleaseFolder = BaseFolder + "ISDA SIMM " + CStr(i) + "\"
7                     Exit Function
8                 End If
9             Next
10        End If

11        WorkbookReleaseFolder = BaseFolder

12        Exit Function
ErrHandler:
13        Throw "#WorkbookReleaseFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveAddinAndExportVBA
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Saves this addin to c:\SolumAddin and saves a time-stamped copy to folder
'             whose location is shown on the Audit sheet ("BackupLocalReleasesTo")
'             also exports the VBA code for Git.
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveAddinAndExportVBA(wb As Excel.Workbook)
          Dim BackupDirectory As String
          Dim BookName As String
          Dim StandardDirectory As String
1         BookName = sStringBetweenStrings(wb.Name, , ".")

          Dim CouldCreateFolder As Boolean
          Dim NowString As String
          Dim Prompt As String
          Dim TimeStamp As Date
          Dim XLAMBackup As String
          Dim XLAMName As String

2         On Error GoTo ErrHandler

3         If Not sFolderIsWritable(wb.Path) Then
4             Throw "Cannot release since " + wb.Path + " is not writable."
5         End If

6         If Not CanMakeReadWrite(wb) Then
7             Throw "Cannot release since attempt to set file access to ReadWrite failed. Use file explorer to check file permissions for " & wb.FullName + vbLf + "You need to have 'Full Control' of this file."
8         End If

9         If wb.VBProject.Protection = 1 Then        'vbext_pp_locked
10            MsgBoxPlus "Please unlock the VBA code of " & gAddinName
11            Application.SendKeys "%{F11}%W{UP}{RETURN}"
12            Exit Sub
13        End If

14        CleanMe

15        TimeStamp = Now()
16        NowString = Format$(TimeStamp, "yyyy-mm-dd hh-mm-ss")

17        BackupDirectory = DataFromAuditSheet("BackupLocalReleasesTo")
18        BackupDirectory = Replace(BackupDirectory, "<BOOKNAME>", sStringBetweenStrings(wb.Name, , "."))

19        If Not sFolderExists(BackupDirectory) Then Throw "BackupDirectory '" & BackupDirectory & "' not found"

20        BackupDirectory = BackupDirectory + NowString + "\"
21        CouldCreateFolder = Not sIsErrorString(sCreateFolder(BackupDirectory))

22        If Not CouldCreateFolder Then
23            Prompt = "Warning. Could not create backup folder " + vbLf + BackupDirectory + vbLf + vbLf + "Release will proceed but without backing up the addin itself and its modules to such folder."
24            MsgBoxPlus Prompt, vbExclamation
25        End If

26        StandardDirectory = DataFromAuditSheet("LocalReleaseFolder")
27        If Not sFolderExists(StandardDirectory) Then Throw "StandardDirectory '" & StandardDirectory & "' not found"

28        XLAMName = StandardDirectory + BookName + ".xlam"
29        XLAMBackup = BackupDirectory + BookName + ".xlam"
          'PGS 21 Oct 2013
          'This line looks unnecessarily restrictive, but I can't get the .SaveAs method to work reliably to save as an XLAM file
          'it silently saves a copy of the workbook with a gibberish name :-(. Hence have to use .Save and that means we must
          'start the release procedure with the document already saved in the right place.
30        If Not FileNamesEquivalent(wb.FullName, LCase$(XLAMName)) Then Throw "Workbook must be saved as:" + vbLf + _
              XLAMName + vbLf + "before the release script is run." + vbLf + "It is currently saved as:" + vbLf + wb.FullName

          'See comments in SaveAddin
31        SaveAddin wb

32        If CouldCreateFolder Then
              Dim CopyRes
33            CopyRes = sFileCopy(XLAMName, XLAMBackup)
34            If sIsErrorString(CopyRes) Then
35                Throw "Failed to copy '" + XLAMName + "' to '" + XLAMBackup + "' " + CopyRes
36            End If
37        End If

          'CleanMe empties out the Range PointerToRibbon, so re-populate it
38        With shAudit.Range("PointerToRibbon")
39            .Value = ObjPtr(g_rbxIRibbonUI)
40        End With

          'Write out the VBA code as well...
41        ExportModules wb, WorkbookToLocalGitFolder(wb), False      'for Git
42        If CouldCreateFolder Then
43            ExportModules wb, BackupDirectory, True
44        End If
45        Exit Sub
ErrHandler:
46        Throw "#SaveAddinAndExportVBA (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation
End Sub

Private Function CanMakeReadWrite(wb As Excel.Workbook) As Boolean
1         On Error GoTo ErrHandler
2         If wb.ReadOnly Then
3             wb.ChangeFileAccess xlReadWrite
4         End If
5         CanMakeReadWrite = Not (wb.ReadOnly)
6         Exit Function
ErrHandler:
7         CanMakeReadWrite = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CleanMe
' Author    : Philip Swannell
' Date      : 01-Mar-2016
' Purpose   : Put this workbook in a good state for release
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CleanMe()

1         On Error GoTo ErrHandler

2         shAudit.Calculate
3         RegisterFunctionsWithFunctionWizard vbNullString         'Make sure we don't release if registration of functions is not working, e.g. because of newly added text breaking the 255 character limit.
4         RefreshIntellisenseSheet shHelp.Range("TheData"), shIntellisense
5         UninstallIntellisense
6         InstallIntellisense

7         CleanUpUndoBuffer shUndo
8         CleanUpUndoBuffer shUndo2
          'Check Help sheet is in good sheet
9         shHelp.Calculate
          Dim c As Range
10        For Each c In shHelp.Range("TheData").Cells
11            If Not IsEmpty(c.Value) Then
12                If VarType(c.Value) <> vbString Then
13                    If Not (IsNumber(c.Value) And c.Column = 6) Then
14                        Throw "Detected non string in Help sheet at cell " + AddressND(c) + ". This needs to be fixed before release."
15                    End If
16                End If
17            End If
18        Next c

19        shAudit.Range("Headers").Cells(2, 2).Value = Date
20        shAudit.Range("Headers").Cells(2, 3).Value = CDbl(Now() - Date)

21        RefreshRibbon        'ensures g_rbxIRibbonUI exists
22        With shAudit.Range("PointerToRibbon")
23            .Locked = False
24            .ClearContents        'ensure that version saved does not have misleading information as to location of pointer to ribbon
25        End With

26        shAudit.Calculate

27        Exit Sub
ErrHandler:
28        Throw "#CleanMe (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FileNamesEquivalent
' Author    : Philip Swannell
' Date      : 21-Apr-2015
' Purpose   : Shameful bodge function to cope with the fact that file.FullName sometimes
'             returns a windows file name for the file on my local C drive and sometimes a
'             URL for some microsoft server in the cloud. This function only works on my PC at Solum!
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileNamesEquivalent(ByVal FileName1 As String, ByVal FileName2 As String) As Boolean
1         If LCase$(FileName1) = LCase$(FileName2) Then
2             FileNamesEquivalent = True
3             Exit Function
4         End If

          'EquivA and EquivB must be lower case and use forward slashes
          Const EquivA = "https://solumfinancial-my.sharepoint.com/personal/philip_swannell_solum-financial_com/documents"
          Const EquivB = "c:/users/philip/onedrive - solum financial limited"

5         FileName1 = LCase$(Replace(FileName1, "\", "/"))
6         FileName2 = LCase$(Replace(FileName2, "\", "/"))
7         FileName1 = Replace(FileName1, EquivB, EquivA)
8         FileName2 = Replace(FileName2, EquivB, EquivA)

9         FileNamesEquivalent = FileName1 = FileName2

10        On Error GoTo ErrHandler

11        Exit Function
ErrHandler:
12        Throw "#FileNamesEquivalent (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: sCellContentsFromFileOnDisk
' Purpose: Returns the contents of a cell on a sheet in a workbook that's saved on disk without opening the workbook
' Parameter FileName (String): Full file name including path
' Parameter SheetName (String): the name of the worksheet
' Parameter CellAddress (String): The A1-style cell address
' Author: Philip Swannell
' Date: 05-Dec-2017
'Note (12 June 2018) Should investigate using Ron de Bruin's method http://www.rondebruin.nl/win/s3/win024.htm
' -----------------------------------------------------------------------------------------------------------------------
Function sCellContentsFromFileOnDisk(FileName As String, SheetName As String, CellAddress As String)
Attribute sCellContentsFromFileOnDisk.VB_Description = "Returns the value of a cell in a workbook without opening the workbook."
Attribute sCellContentsFromFileOnDisk.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim file As String
          Dim FullReference As String
          Dim Path As String
          Dim R1C1Address As String

1         On Error GoTo ErrHandler
2         Path = sSplitPath(FileName, False)
3         Path = Replace(Path, "'", "''")
4         file = sSplitPath(FileName, True)
5         file = Replace(file, "'", "''")
6         R1C1Address = Application.ConvertFormula(CellAddress, xlA1, xlR1C1, True, shAudit.Range("A1"))
7         FullReference = "'" & Path & "\[" & file & "]" & SheetName & "'!" & R1C1Address
8         sCellContentsFromFileOnDisk = ExecuteExcel4Macro2(FullReference)
9         Exit Function
ErrHandler:
10        sCellContentsFromFileOnDisk = "#sCellContentsFromFileOnDisk (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: CheckForOverwrite
' Purpose: Check whether releasing a workbook is likely to overwrite other people's work
' Author: Philip Swannell
' Date: 21-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Sub CheckForOverwrite(WbToRelease As Excel.Workbook, isAddin As Boolean, Optional CheckOnly As Boolean = False)

          Dim strFile As String
          Dim strPath As String
          Const shName = "Audit"
          Dim FailReason As String
          Dim shortFailReason As String

          Dim AllowOverride As Boolean
          Dim cn_Comment As Variant
          Dim cn_Version As Variant
          Dim CommentCell As Range
          Dim GoodToRelease As Boolean
          Dim Headers
          Dim ReleasedComment As Variant
          Dim ReleasedVersion As Variant
          Dim ReleaseTime As Variant
          Dim shAudit As Worksheet
          Dim strReleaseTime As String
          Dim ThisComment As String
          Dim ThisVersion As Long
          Dim VersionCell As Range

1         On Error GoTo ErrHandler

2         If Not IsInCollection(WbToRelease.Worksheets, shName) Then Throw "There must be an " + shName + " sheet in the workbook to be released"
3         Set shAudit = WbToRelease.Worksheets(shName)
4         Headers = sArrayTranspose(RangeFromSheet(shAudit, "Headers").Value)
5         cn_Version = sMatch("Version", Headers)
6         If Not IsNumber(cn_Version) Then Throw "Cannot find text ""Version"" in Range ""Headers"" of sheet """ + shName + """"
7         Set VersionCell = shAudit.Range("Headers").Cells(2, cn_Version)
8         cn_Comment = sMatch("Comment", Headers)
9         If Not IsNumber(cn_Comment) Then Throw "Cannot find text ""Comment"" in Range ""Headers"" of sheet """ + shName + """"
10        Set CommentCell = shAudit.Range("Headers").Cells(2, cn_Comment)

11        If isAddin Then
12            strPath = DataFromAuditSheet("NetworkReleaseFolder") + "Addins\"
13        Else
14            strPath = WorkbookReleaseFolder(WbToRelease.Name)
15        End If
16        strFile = WbToRelease.Name

17        If Not sFileExists(strPath + strFile) Then
18            Exit Sub
19        Else
20            ReleaseTime = sFileLastModifiedDate(strPath + strFile)
21            strReleaseTime = Format$(ReleaseTime, "dd-mmm-yyyy hh:mm")
22        End If

23        ReleasedComment = sCellContentsFromFileOnDisk(strPath + strFile, shName, CommentCell.address)
24        ReleasedVersion = sCellContentsFromFileOnDisk(strPath + strFile, shName, VersionCell.address)

25        ThisVersion = VersionCell.Value
26        ThisComment = CommentCell.Value

27        GoodToRelease = True

28        If sEquals(ThisVersion, ReleasedVersion) Then
29            If sEquals(ThisComment, ReleasedComment) Then
30                FailReason = "The version number (" + Format$(ThisVersion, "#,###") + ") and release Comment on the Audit sheet are the same as those for the most recent release in" + _
                      vbLf + "'" + strPath + "' which is dated " + strReleaseTime + _
                      vbLf + vbLf + "Please add a new line and release comment to the Audit sheet of this workbook."
31                GoodToRelease = False
32                shortFailReason = "the version number of this copy is the same as that of the latest release."
33                If Not CheckOnly Then Throw FailReason, True
34            Else
35                GoodToRelease = False
36                FailReason = "Version " + Format$(ThisVersion, "#,###") + " has already been released with Comment:" + vbLf + vbLf + ReleasedComment + vbLf + vbLf + _
                      "Please check that this release is not overwriting other people's work"
37                shortFailReason = "it appears that this copy does not include the most recent changes in the released version."
38                AllowOverride = True
39                If Not CheckOnly Then
                      Dim MsgBoxRes As VbMsgBoxResult
40                    MsgBoxRes = MsgBoxPlus(FailReason, vbOKCancel + vbDefaultButton2 + vbExclamation, , "Release anyway")
41                    If MsgBoxRes <> vbOK Then Throw "Release aborted", True
42                End If

43            End If
44        End If

45        If (ThisVersion < ReleasedVersion) Then
46            GoodToRelease = False
47            FailReason = "Version " + Format$(ReleasedVersion, "#,###") + " has already been released with Comment:" + vbLf + vbLf + ReleasedComment + vbLf + vbLf + _
                  "So releasing this workbook would be likely to overwrite work done by you or other team members"
48            shortFailReason = "it appears that this copy does not include the most recent changes in the released version."
49            AllowOverride = True
50            If Not CheckOnly Then
51                MsgBoxRes = MsgBoxPlus(FailReason, vbOKCancel + vbDefaultButton2 + vbExclamation, , "Release anyway")
52                If MsgBoxRes <> vbOK Then Throw "Release aborted", True
53            End If

54        End If

55        If ThisVersion > ReleasedVersion Then
              Dim LookedUpComment As String
56            LookedUpComment = sVLookup(ReleasedVersion, sExpandDown(shAudit.Range("Headers")), "Comment", "Version")
57            If LookedUpComment <> ReleasedComment Then
58                GoodToRelease = False
59                FailReason = "The most recent release of this workbook (version " + Format$(ReleasedVersion, "#,###") + " dated " + strReleaseTime + ") had the following release comment:" + vbLf + vbLf + _
                      ReleasedComment + vbLf + vbLf + _
                      "However the Audit sheet of this workbook has a different comment associated with that release number, namely:" + vbLf + vbLf + _
                      LookedUpComment + vbLf + vbLf + _
                      "So releasing this workbook would be likely to overwrite work done by you or other team members"
60                shortFailReason = "it appears that this copy does not include the most recent changes in the released version."
61                AllowOverride = True
62                If Not CheckOnly Then
63                    MsgBoxRes = MsgBoxPlus(FailReason, vbOKCancel + vbDefaultButton2 + vbExclamation, , "Release anyway")
64                    If MsgBoxRes <> vbOK Then Throw "Release aborted", True
65                End If
66            End If
67        End If

68        If CheckOnly Then
              Dim CheckPrompt As String
              Dim PromptArray
69            PromptArray = sArrayStack("This copy", vbNullString, _
                  "Version", Format$(ThisVersion, "#,###"), _
                  "Comment", ThisComment, _
                  "Path", WbToRelease.Path, _
                  "Latest release", vbNullString, _
                  "Version", Format$(ReleasedVersion, "#,###"), _
                  "Comment", ReleasedComment, _
                  "Release time", strReleaseTime + ", " + sDescribeTime(Now - ReleaseTime), _
                  "Path", strPath)
70            PromptArray = sReshape(PromptArray, sNRows(PromptArray) / 2, 2)
71            PromptArray = sJustifyArrayOfStrings(PromptArray, "Segoe UI", 9, vbTab)
72            CheckPrompt = sConcatenateStrings(PromptArray, vbLf)
73            CheckPrompt = "Comparison of this copy of " + WbToRelease.Name + " to the latest release:" + vbLf + CheckPrompt
74            If GoodToRelease Then
75                CheckPrompt = CheckPrompt + vbLf + vbLf + "The workbook could be released."
76            ElseIf AllowOverride Then
77                CheckPrompt = CheckPrompt + vbLf + vbLf + "WARNING" + vbLf + shortFailReason + " But release will still be possible"
78            Else
79                CheckPrompt = CheckPrompt + vbLf + vbLf + "WARNING" + vbLf + "Workbook release would fail because " + shortFailReason
80            End If
81            MsgBoxPlus CheckPrompt, IIf(GoodToRelease, vbInformation, vbExclamation), , , , , , 400
82            Exit Sub
83        End If

84        Exit Sub
ErrHandler:
85        Throw "#CheckForOverwrite (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: AddReleaseCommentToMRU
' Purpose: Utility to make it easy to use previous comments when we release a workbook
' Parameter Comment (String):
' Author: Philip Swannell
' Date: 04-Dec-2017
' -----------------------------------------------------------------------------------------------------------------------
Sub AddReleaseCommentToMRU(Comment As String)
          Dim Comments As Variant
          Const MaxNumComments = 30

1         On Error GoTo ErrHandler
2         Comments = GetSetting(gAddinName, "ReleaseComments", "MRU", "Not found")
3         If Comments <> "Not found" Then
4             Comments = sArrayStack(Comment, sParseArrayString(CStr(Comments)))
5             Comments = sRemoveDuplicates(Comments, False, False)
6             If sNRows(Comments) > MaxNumComments Then
7                 Comments = sSubArray(Comments, 1, 1, MaxNumComments)
8             End If
9         Else
10            Comments = Comment
11        End If
12        SaveSetting gAddinName, "ReleaseComments", "MRU", sMakeArrayString(Comments)

13        Exit Sub
ErrHandler:
14        Throw "#AddReleaseCommentToMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function ReleaseCommentFromAuditSheet(ws As Worksheet)
          Dim ColNo As Variant
          Dim R As Range
1         On Error GoTo ErrHandler
2         Set R = RangeFromSheet(ws, "Headers")
3         ColNo = sMatch("Comment", sArrayTranspose(R.Value))
4         If sIsErrorString(ColNo) Then Throw "Cannot find header ""Comment"" in range ""Headers"" on sheet " + ws.Parent.Name + "!" + ws.Name
5         ReleaseCommentFromAuditSheet = R.Cells(2, CLng(ColNo)).Value

6         Exit Function
ErrHandler:
7         Throw "#ReleaseCommentFromAuditSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub AddReleaseCommentFromAuditSheetToMRU(ws As Worksheet)
1         AddReleaseCommentToMRU ReleaseCommentFromAuditSheet(ws)
2         Exit Sub
ErrHandler:
3         Throw "#AddReleaseCommentFromAuditSheetToMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: SaveAddin
' Purpose:  Saves an addin to its current location. This turns out to be frustratingly difficult.
' Things that can go wrong:
' 1) using .SaveAs rather than .Save can produce a randomly-named file in the correct folder
' 2) using Save can save an xlsm to the documents folder (or perhaps the folder given by CurDir$) rather than save an xlam to the correct folder.
' Top Tips: Ensure that logged on as administrator and that the folder to be saved to is "owned" by the same user.
' This article helped:
' https://social.technet.microsoft.com/Forums/windows/en-US/f0d5d4c0-c11c-435a-aa2d-83388e3bd2d2/i-am-an-administrator-but-i-keep-getting-errors-like-quotyou-require-permission-from?forum=w7itprosecurity
' Parameter wb (Workbook):
' Author: Philip Swannell
' Date: 15/10/2018
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SaveAddin(wb As Excel.Workbook)
          Dim NewPath As String
          Dim origIsAddin As Boolean
          Dim origPath As String
          Dim origTime

1         On Error GoTo ErrHandler
          Dim XSH As clsExcelStateHandler
2         Set XSH = CreateExcelStateHandler(, , False)
3         origPath = wb.FullName
4         origIsAddin = wb.isAddin
5         origTime = sFileInfo(origPath, "LastModifiedDate")
6         If Not origIsAddin Then wb.isAddin = True
7         wb.Save
8         NewPath = wb.FullName
9         If Not origIsAddin Then wb.isAddin = False
10        If NewPath <> origPath Then
11            Throw "Error: saving the workbook changed its location from:" + vbLf + "'" + origPath + "'" + vbLf + "to:" + vbLf + "'" + NewPath + "'"
12        End If
13        If sFileInfo(origPath, "LastModifiedDate") = origTime Then
14            Throw "Assertion failed: Call to .Save method for workbook '" + origPath + "' did not change its last modified date"
15        End If
16        Exit Sub
ErrHandler:
17        Throw "#SaveAddin (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub SaveMe()
1         SaveAddin ThisWorkbook
End Sub
