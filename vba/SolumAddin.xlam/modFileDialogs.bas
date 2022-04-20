Attribute VB_Name = "modFileDialogs"
Option Explicit
Private Const MAXMRUs = 10        'Determines the number of "Most Recently Used Files" displayed by GetOpenFilenameWrap
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetSaveAsFilenameWrap
' Author    : Philip Swannell
' Date      : 13-Jul-2015
' Purpose   : Wrap to Application.GetSaveAsFilename but can pass registry key so that we
'             "remember" where various types of files reside.
' -----------------------------------------------------------------------------------------------------------------------
Function GetSaveAsFilenameWrap(RegKey As String, Optional InitialFileName, Optional FileFilter As Variant, Optional FilterIndex, Optional Title, Optional ButtonText)
          Dim CurrentDirectory As String
          Dim NewDirectory
          Dim Prompt As String
          Dim Res As Variant
          Dim SavedDirectory As String

1         On Error GoTo ErrHandler

2         If TypeName(Application.Caller) = "Range" Then
3             GetSaveAsFilenameWrap = "#Function GetSaveAsFilenameWrap cannot be called from a spreadsheet!"        'Calling from sheet will crash Excel
4             Exit Function
5         End If

6         CurrentDirectory = CurDir$()
7         SavedDirectory = GetSetting(gAddinName, "FileLocations", RegKey)

8         If SavedDirectory <> vbNullString Then
9             On Error Resume Next        'directory may no longer exist or no longer be accessible
10            ChDir SavedDirectory
11            On Error GoTo ErrHandler
12        End If

TryAgain:
13        Res = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, Title, ButtonText)

14        If VarType(Res) <> vbBoolean Then
15            If sFileExists(Res) Then
16                Prompt = Res + " already exists." + vbLf + "Do you want to replace it?"
17                If MsgBoxPlus(Prompt, vbYesNo + vbExclamation + vbDefaultButton2, Title, , , , , 400) <> vbYes Then GoTo TryAgain
18            End If
19        End If

20        If VarType(Res) <> vbBoolean Then
21            If VarType(Res) = vbString Then
22                NewDirectory = sSplitPath(CStr(Res), False)
23                AddFileToMRU RegKey, CStr(Res)
24            ElseIf VarType(Res) = vbArray Then
25                NewDirectory = sSplitPath(CStr(Res(1)), False)
26                AddFileToMRU RegKey, CStr(Res(1))
27            End If
28            SaveSetting gAddinName, "FileLocations", RegKey, NewDirectory
29        End If

30        ChDir CurrentDirectory

31        GetSaveAsFilenameWrap = Res

32        Exit Function
ErrHandler:
33        Throw "#GetSaveAsFilenameWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MoveMRUFileList
' Author    : Philip
' Date      : 13-Jun-2017
' Purpose   : If we want to change the RegKey used for storing folder locations and mru file lists
'             this method will do the move from the OldRegKey to NewRegKey but only if there is
'             data stored at the old key and no data stotred at the new key
' -----------------------------------------------------------------------------------------------------------------------
Sub MoveMRUFileList(OldRegKey As String, NewRegKey As String)
          Dim NewFolder As String
          Dim NewMRU As String
          Dim OldFolder As String
          Dim OldMRU As String
          Const nf = "Not found"

1         On Error GoTo ErrHandler
2         OldFolder = GetSetting(gAddinName, "FileLocations", OldRegKey, nf)
3         OldMRU = GetSetting(gAddinName, "FileLocations", OldRegKey + "MRU", nf)
4         NewFolder = GetSetting(gAddinName, "FileLocations", NewRegKey, nf)
5         NewMRU = GetSetting(gAddinName, "FileLocations", NewRegKey + "MRU", nf)

6         If OldFolder <> nf And NewFolder = nf And OldMRU <> nf And NewMRU = nf Then
7             SaveSetting gAddinName, "FileLocations", NewRegKey, OldFolder
8             SaveSetting gAddinName, "FileLocations", NewRegKey + "MRU", OldMRU
9             DeleteSetting gAddinName, "FileLocations", OldRegKey
10            DeleteSetting gAddinName, "FileLocations", OldRegKey + "MRU"
11        End If

12        Exit Sub
ErrHandler:
13        Throw "#MoveMRUFileList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetMRUList
' Author    : Philip Swannell
' Date      : 08-Mar-2016
' Purpose   : For use in conjunction with GetOpenFilenameWrap and ShowCommandBarPopup
'             to implement "Most recently used file" lists.
' -----------------------------------------------------------------------------------------------------------------------
Sub GetMRUList(RegKey As String, ByRef FileNames, ByRef TheChoices, ByRef FaceIDs, ByRef EnableFlags, FolderMode As Boolean)
          Dim i As Long
          Dim MRUFilesAbbreviated As Variant

1         On Error GoTo ErrHandler
2         FileNames = GetSetting(gAddinName, "FileLocations", RegKey + "MRU", "Not found")
3         If FileNames = "Not found" Then
4             FileNames = Empty: TheChoices = Empty: FaceIDs = Empty: EnableFlags = Empty
5             Exit Sub
6         End If

7         FileNames = sRemoveDuplicates(sFileMappedToUNC(sTokeniseString(CStr(FileNames), vbTab)))

8         If sNRows(FileNames) > MAXMRUs Then
9             FileNames = sSubArray(FileNames, 1, 1, MAXMRUs)
10        End If
11        Force2DArray FileNames
12        MRUFilesAbbreviated = FileNames
13        For i = 1 To sNRows(MRUFilesAbbreviated)
14            If i <= 9 Then
15                MRUFilesAbbreviated(i, 1) = AbbreviateForCommandBar("&" & CStr(i) + " " + CStr(MRUFilesAbbreviated(i, 1)))
16            ElseIf i = 10 Then
17                MRUFilesAbbreviated(i, 1) = AbbreviateForCommandBar("1&0" + " " + CStr(MRUFilesAbbreviated(i, 1)))
18            Else
19                MRUFilesAbbreviated(i, 1) = AbbreviateForCommandBar(CStr(MRUFilesAbbreviated(i, 1)))
20            End If
21        Next i

22        TheChoices = sArrayStack(MRUFilesAbbreviated, "--&Browse" & IIf(FolderMode, " for folder...", "..."), "--Clear &History")
23        FaceIDs = sReshape(0, sNRows(TheChoices), 1)
24        For i = 1 To SafeMin(10, sNRows(TheChoices) - 2)
25            FaceIDs(i, 1) = 70 + i Mod 10
26        Next i
27        FaceIDs(sNRows(TheChoices) - 1, 1) = 23    'Browse
28        FaceIDs(sNRows(TheChoices), 1) = 358  'Delete
29        EnableFlags = sReshape(True, sNRows(TheChoices), 1)

30        Exit Sub
ErrHandler:
31        Throw "#GetMRUList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetOpenFilenameWrap
' Author    : Philip Swannell
' Date      : 20-Nov-2017
' Purpose   : Wrap to Application.GetOpenFileName but can pass registry
'             key so that we "remember" where various types of files reside.
'             If WithMRU is TRUE then we first show a command-bar listing "Most Recently Used"
'             files with a "Browse..." option at the bottom.
'             If the user cancels out of the dialog then the return is FALSE otherwise
'             If MultiSelect is TRUE then the return is a 1-dimensional array of file names; if MultiSelect is FALSE then the return is a string
' -----------------------------------------------------------------------------------------------------------------------
Function GetOpenFilenameWrap(RegKey As String, Optional FileFilter As Variant = "All Files (*.*),*.*", Optional FilterIndex = 1, _
        Optional Title, Optional ButtonText, Optional MultiSelect As Boolean, Optional WithMRU As Boolean, Optional AnchorObject As Object)

          Dim MRUFiles As Variant
          Dim NewDirectory
          Dim Res As Variant
          Dim SavedDirectory As String

1         On Error GoTo ErrHandler

2         If TypeName(Application.Caller) = "Range" Then
3             GetOpenFilenameWrap = "Function GetOpenFilenameWrap cannot be called from a spreadsheet!"        'Calling from sheet can crash Excel
4             Exit Function
5         End If

6         If WithMRU Then
              Dim ChosenFile As String
              Dim EnableFlags
              Dim FaceIDs
              Dim TheChoices
7             GetMRUList RegKey, MRUFiles, TheChoices, FaceIDs, EnableFlags, False
8             If Not IsEmpty(MRUFiles) Then
9                 Res = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , AnchorObject, True)
10                If Res = 0 Then
11                    GetOpenFilenameWrap = False
12                    Exit Function
13                ElseIf Res = sNRows(TheChoices) Then
14                    RemoveFileFromMRU RegKey, "All"
15                    Exit Function
16                ElseIf Res = sNRows(TheChoices) - 1 Then
                      'Browse
17                Else
18                    ChosenFile = MRUFiles(Res, 1)
19                    GetOpenFilenameWrap = sFileMappedToUNC(ChosenFile)
20                    AddFileToMRU RegKey, ChosenFile
21                    Exit Function
22                End If
23            End If
24        End If

25        SavedDirectory = GetSetting(gAddinName, "FileLocations", RegKey)
26        If SavedDirectory <> vbNullString Then If Right$(SavedDirectory, 1) <> "\" Then SavedDirectory = SavedDirectory + "\"

          Dim OldCurDir As String
27        OldCurDir = CurDir$()

28        If (SavedDirectory <> vbNullString) Then
29            On Error Resume Next
30            ChDrive SavedDirectory    'operates on first letter
31            ChDir SavedDirectory
32            On Error GoTo ErrHandler
33        End If

34        Res = Application.GetOpenFileName(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)

35        If VarType(Res) <> vbBoolean Then
36            If VarType(Res) = vbString Then
37                NewDirectory = sSplitPath(CStr(Res), False)
38            ElseIf VarType(Res) = vbArray Then
39                NewDirectory = sSplitPath(CStr(Res(1)), False)
40            ElseIf (IsArray(Res)) Then
41                NewDirectory = sSplitPath(CStr(Res(1)), False)
42            End If
43            SaveSetting gAddinName, "FileLocations", RegKey, NewDirectory
44            Res = sFileMappedToUNC(Res)
45        End If

46        GetOpenFilenameWrap = Res
47        On Error Resume Next
48        ChDrive OldCurDir
49        ChDir OldCurDir    'leave the current directory unchanged, though not if that's necessary
50        On Error GoTo ErrHandler

51        If VarType(Res) >= vbArray Then
52            If sNRows(Res) = 1 Then
53                AddFileToMRU RegKey, CStr(Res(1, 1))
54            End If
55        ElseIf VarType(Res) = vbString Then
56            AddFileToMRU RegKey, CStr(Res)
57        End If

58        Exit Function
ErrHandler:
59        Throw "#GetOpenFilenameWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddFileToMRU
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Encapsulates adding a file to the MRU list held in the Registry
' -----------------------------------------------------------------------------------------------------------------------
Sub AddFileToMRU(RegKey As String, FileName As String)
          Dim MRUFiles As Variant

1         On Error GoTo ErrHandler
2         MRUFiles = GetSetting(gAddinName, "FileLocations", RegKey + "MRU", "Not found")
3         If MRUFiles <> "Not found" Then
4             MRUFiles = sArrayStack(FileName, sTokeniseString(CStr(MRUFiles), vbTab))
5             MRUFiles = sRemoveDuplicates(MRUFiles, False, False)
6             If sNRows(MRUFiles) > MAXMRUs Then
7                 MRUFiles = sSubArray(MRUFiles, 1, 1, MAXMRUs)
8             End If
9         Else
10            MRUFiles = FileName
11        End If
12        SaveSetting gAddinName, "FileLocations", RegKey + "MRU", sConcatenateStrings(MRUFiles, vbTab)

13        Exit Sub
ErrHandler:
14        Throw "#AddFileToMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RemoveFileFromMRU
' Author    : Philip Swannell
' Date      : 14-Mar-2016
' Purpose   : Encapsulates removing a file from an MRU list. Pass FileName as "All" to remove all
' -----------------------------------------------------------------------------------------------------------------------
Sub RemoveFileFromMRU(RegKey As String, FileName As String)
          Dim ChooseVector As Variant
          Dim MRUFiles As Variant
1         On Error GoTo ErrHandler
2         MRUFiles = GetSetting(gAddinName, "FileLocations", RegKey + "MRU", "Not found")
3         If MRUFiles <> "Not found" Then
4             MRUFiles = sTokeniseString(CStr(MRUFiles), vbTab)
5             ChooseVector = sArrayNot(sArrayEquals(MRUFiles, FileName))
6             If sArraysIdentical(True, sColumnOr(ChooseVector)) And FileName <> "All" Then
7                 MRUFiles = sMChoose(MRUFiles, ChooseVector)
8                 SaveSetting gAddinName, "FileLocations", RegKey + "MRU", sConcatenateStrings(MRUFiles, vbTab)
9             Else
10                DeleteSetting gAddinName, "FileLocations", RegKey + "MRU"
11            End If
12        End If
13        Exit Sub
ErrHandler:
14        Throw "#RemoveFileFromMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderPicker
' Author    : Philip Swannell
' Date      : 10-Feb-2016
' Purpose   : Wrap to Application.FileDialog(msoFileDialogFolderPicker).
'             Returns False if the user cancels out of the dialog, else returns the name of the chosen folder (without trailing backslash)
'             Saves the chosen folder to the registry so that dialog can "remember" the users last chosen location
' -----------------------------------------------------------------------------------------------------------------------
Function FolderPicker(Optional ByVal InitialFolder As String, Optional ButtonName As String = "OK", _
        Optional Title As String = "Select Folder", Optional RegKey As String, Optional WithMRU As Boolean, _
        Optional AnchorObject As Object)

1         On Error GoTo ErrHandler

2         If WithMRU Then
              Dim ChosenFolder As String
              Dim EnableFlags
              Dim FaceIDs
              Dim MRUFiles
              Dim Res
              Dim TheChoices
3             GetMRUList RegKey, MRUFiles, TheChoices, FaceIDs, EnableFlags, True
4             If Not IsEmpty(MRUFiles) Then
5                 Res = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , AnchorObject, True)
6                 If Res = 0 Then
7                     FolderPicker = False
8                     Exit Function
9                 ElseIf Res = sNRows(TheChoices) Then
10                    RemoveFileFromMRU RegKey, "All"
11                    Exit Function
12                ElseIf Res = sNRows(TheChoices) - 1 Then
                      'Browse
13                Else
14                    ChosenFolder = sFileMappedToUNC(MRUFiles(Res, 1))
15                    FolderPicker = ChosenFolder
16                    AddFileToMRU RegKey, ChosenFolder
17                    SaveSetting gAddinName, "FolderLocations", RegKey, ChosenFolder
18                    Exit Function
19                End If
20            End If
21        End If

22        With Application.FileDialog(msoFileDialogFolderPicker)
23            If .Title <> vbNullString Then
24                .Title = Title
25            End If
26            If ButtonName <> vbNullString Then
27                .ButtonName = ButtonName
28            End If
29            .InitialView = msoFileDialogViewList
30            .AllowMultiSelect = False        'Setting .AllowMultiSelect to True seems to not work
31            If InitialFolder = vbNullString Then
32                If RegKey <> vbNullString Then
33                    InitialFolder = GetSetting(gAddinName, "FolderLocations", RegKey, vbNullString)
34                End If
35            End If
36            If InitialFolder <> vbNullString Then
37                If Right$(InitialFolder, 1) <> "\" Then InitialFolder = InitialFolder + "\"
38                If sFolderExists(InitialFolder) Then
39                    .InitialFileName = InitialFolder
40                End If
41            End If

42            If .Show Then
43                ChosenFolder = .SelectedItems(1)
44                ChosenFolder = sFileMappedToUNC(ChosenFolder)
45                If RegKey <> vbNullString Then
46                    SaveSetting gAddinName, "FolderLocations", RegKey, ChosenFolder
47                End If
48                AddFileToMRU RegKey, ChosenFolder
49                FolderPicker = ChosenFolder
50            Else
51                FolderPicker = False
52            End If
53        End With
54        Exit Function
ErrHandler:
55        FolderPicker = "#FolderPicker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
