Attribute VB_Name = "modMe"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modMe
' Author    : Philip
' Date      : 24-Dec-2015
' Purpose   : Code relating to SolumAddin itself - e.g. releasing new versions, exporting the VBA code etc.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestHelpSheetData
' Author     : Philip Swannell
' Date       : 24-Jul-2019
' Purpose    : Add-hoc code to test if the contents of column C of the Help sheet is correct. Populates a column of a
'              worksheet in a separate instance of Excel with what column C of the Help sheet "should" contain.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestHelpSheetData()
          Dim allFunctions As Variant
          Dim AllXLs As Collection
          Dim FunctionName As String
          Dim i As Long
          Dim OtherExcel As Excel.Application
          Dim t1 As Double
          Dim wb As Workbook
          Const FirstFunction = "sAll" ' <- change if we document a function that's earlier in the alphabet

1         On Error GoTo ErrHandler
2         GetExcelInstances AllXLs

3         If AllXLs.Count = 1 Then Throw "Please Launch a second instance of Excel (use Alt key)", True
4         Set OtherExcel = AllXLs(2)

5         Set wb = OtherExcel.Workbooks.Add
6         AppActivate OtherExcel.caption

7         allFunctions = shHelp.Range("TheData").Columns(1).Value
8         allFunctions = sSubArray(allFunctions, sMatch(FirstFunction, allFunctions))

9         For i = 1 To sNRows(allFunctions)
10            FunctionName = allFunctions(i, 1)
11            Application.SendKeys "=" & FunctionName & "^+A{F2}{Home}{DELETE}{ENTER}{DOWN}"
12            t1 = sElapsedTime
13            While sElapsedTime < t1 + 0.25
14                DoEvents
15            Wend
16        Next

17        g wb.Worksheets(1).UsedRange.Value

18        Exit Sub
ErrHandler:
19        SomethingWentWrong "#TestHelpSheetData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ThisAddinPreferences
' Author     : Philip Swannell
' Date       : 17-Jan-2018
' Purpose    : Allow the user control over how SolumAddin changes Excel's Calculation and AutoSave behaviour
' -----------------------------------------------------------------------------------------------------------------------
Sub ThisAddinPreferences()
          Dim AutoSaveRelevant As Boolean
          Dim SUH As clsScreenUpdateHandler
          Dim wb As Excel.Workbook

          Dim CheckBoxText As String
          Dim CheckBoxValue As Boolean
          Dim CurrentChoice
          Dim NewChoice
          Dim TheChoices
          Dim Title As String
          Dim TopText

          Const chCalc1 = "Nothing"
          Const chCalc1b = "&Let Excel decide"

          Const chCalc2 = "xlCalculationAutomatic"
          Const chCalc2b = "&Automatic"

          Const chCalc3 = "xlCalculationSemiautomatic"
          Const chCalc3b = "Automatic e&xcept for data tables"

          Const chCalc4 = "xlCalculationManual"
          Const chCalc4b = "&Manual"

          Const chSave1 = "LeaveAlone"
          Const chSave1b = "Let Excel &decide what AutoSave should be"

          Const chSave2 = "SwitchOff"
          Const chSave2b = "Switch AutoSave &OFF"

          Const chSave3 = "SwitchOn"
          Const chSave3b = "Switch AutoSave O&N"
          
          Const chDev1 = "Standard"
          Const chDev1b = "No, I don't edit &VBA code"
          
          Const chDev2 = "Developer"
          Const chDev2b = "Yes, I edit VBA code"
          
          Const chAsk1 = "Ask before pasting values"
          Const chAsk2 = "Paste values without asking"
          
          Const chPID1 = "Never"
          Const chPID2 = "When multiple Excels open"
          Const chPID3 = "Always"
          
1         On Error GoTo ErrHandler

2         Title = gAddinName & " Preferences"
3         CheckBoxText = "Refresh pivot tables with sheet calculation"
4         TopText = sArrayRange("At Excel startup set workbook calculation to:", "When opening workbooks:", "Use Developer Mode?", "When using Ctrl Shift V", "Show Process ID in Excel Caption")
            
5         TheChoices = sArrayRange(sArrayStack(chCalc1b, chCalc2b, chCalc3b, chCalc4b), _
              sArrayStack(chSave1b, chSave2b, chSave3b), sArrayStack(chDev1b, chDev2b), sArrayStack(chAsk1, chAsk2), sArrayStack(chPID1, chPID2, chPID3))

6         AutoSaveRelevant = Val(Application.Version) >= 16

7         Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "Application.Calculation", chCalc1))
              Case LCase$(chCalc1)
8                 CurrentChoice = chCalc1b
9             Case LCase$(chCalc2)
10                CurrentChoice = chCalc2b
11            Case LCase$(chCalc3)
12                CurrentChoice = chCalc3b
13            Case LCase$(chCalc4)
14                CurrentChoice = chCalc4b
15            Case Else
16                CurrentChoice = chCalc1b
17        End Select

18        Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "Application.AutoSave", chSave1))
              Case LCase$(chSave1)
19                CurrentChoice = sArrayRange(CurrentChoice, chSave1b)
20            Case LCase$(chSave2)
21                CurrentChoice = sArrayRange(CurrentChoice, chSave2b)
22            Case LCase$(chSave3)
23                CurrentChoice = sArrayRange(CurrentChoice, chSave3b)
24            Case Else
25                CurrentChoice = sArrayRange(CurrentChoice, chSave1b)
26        End Select

27        Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "DeveloperMode", chDev1))
              Case LCase$(chDev1)
28                CurrentChoice = sArrayRange(CurrentChoice, chDev1b)
29            Case LCase$(chDev2)
30                CurrentChoice = sArrayRange(CurrentChoice, chDev2b)
31            Case Else
32                CurrentChoice = sArrayRange(CurrentChoice, chDev1b)
33        End Select

34        Select Case LCase$(GetSetting(gAddinName, "PasteValues", "AskBeforePasting", "True"))
              Case "true"
35                CurrentChoice = sArrayRange(CurrentChoice, chAsk1)
36            Case Else
37                CurrentChoice = sArrayRange(CurrentChoice, chAsk2)
38        End Select

39        Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID1))
              Case LCase$(chPID1)
40                CurrentChoice = sArrayRange(CurrentChoice, chPID1)
41            Case LCase$(chPID2)
42                CurrentChoice = sArrayRange(CurrentChoice, chPID2)
43            Case LCase$(chPID3)
44                CurrentChoice = sArrayRange(CurrentChoice, chPID3)
45            Case Else
46                CurrentChoice = sArrayRange(CurrentChoice, chPID1)
47        End Select

48        Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "RefreshPivotTablesWithSheetCalculation", "False"))
              Case "true"
49                CheckBoxValue = True
50            Case Else
51                CheckBoxValue = False
52        End Select

53        EnsureAppObjectExists

54        NewChoice = ShowOptionButtonDialog(TheChoices, Title, TopText, CurrentChoice, , True, CheckBoxText, CheckBoxValue)

55        If sArraysIdentical(NewChoice, sReshape(0, 1, sNCols(TheChoices))) Then Exit Sub

56        Set SUH = CreateScreenUpdateHandler()
57        Set wb = Application.Workbooks.Add

58        Select Case NewChoice(1, 1)
              Case 1
59                SaveSetting gAddinName, "InstallInformation", "Application.Calculation", chCalc1
60            Case 2
61                Application.Calculation = xlCalculationAutomatic
62                Application.CalculateBeforeSave = True
63                SaveSetting gAddinName, "InstallInformation", "Application.Calculation", chCalc2
64            Case 3
65                Application.Calculation = xlCalculationSemiautomatic
66                Application.CalculateBeforeSave = True
67                SaveSetting gAddinName, "InstallInformation", "Application.Calculation", chCalc3
68            Case 4
69                Application.Calculation = xlCalculationManual
70                Application.CalculateBeforeSave = False
71                SaveSetting gAddinName, "InstallInformation", "Application.Calculation", chCalc4
72            Case Else
73                Throw "Assertion failed - unexpected value in return from method ShowOptionButtonDialog"
74        End Select
75        wb.Close False

76        Select Case NewChoice(1, 2)
              Case 1
77                SaveSetting gAddinName, "InstallInformation", "Application.AutoSave", chSave1
78            Case 2
79                SaveSetting gAddinName, "InstallInformation", "Application.AutoSave", chSave2
80            Case 3
81                SaveSetting gAddinName, "InstallInformation", "Application.AutoSave", chSave3
82            Case Else
83                Throw "Assertion failed - unexpected value in return from method ShowOptionButtonDialog"
84        End Select

85        Select Case NewChoice(1, 3)
              Case 1
86                SaveSetting gAddinName, "InstallInformation", "DeveloperMode", chDev1
87            Case 2
88                SaveSetting gAddinName, "InstallInformation", "DeveloperMode", chDev2
89            Case Else
90                Throw "Assertion failed - unexpected value in return from method ShowOptionButtonDialog"
91        End Select

92        Select Case NewChoice(1, 4)
              Case 1
93                SaveSetting gAddinName, "PasteValues", "AskBeforePasting", "True"
94            Case 2
95                SaveSetting gAddinName, "PasteValues", "AskBeforePasting", "False"
96            Case Else
97                Throw "Assertion failed - unexpected value in return from method ShowOptionButtonDialog"
98        End Select

99        Select Case NewChoice(1, 5)
              Case 1
100               SaveSetting gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID1
101           Case 2
102               SaveSetting gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID2
103           Case 3
104               SaveSetting gAddinName, "InstallInformation", "ShowPIDInExcelCaption", chPID3
105           Case Else
106               Throw "Assertion failed - unexpected value in return from method ShowOptionButtonDialog"
107       End Select
108       SetApplicationCaptions

109       If CheckBoxValue Then
110           SaveSetting gAddinName, "InstallInformation", "RefreshPivotTablesWithSheetCalculation", "True"
111       Else
112           SaveSetting gAddinName, "InstallInformation", "RefreshPivotTablesWithSheetCalculation", "False"
113       End If

114       Exit Sub
ErrHandler:
115       SomethingWentWrong "#ThisAddinPreferences (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : WorkbookToLocalGitFolder
' Author    : Philip Swannell
' Date      : 18-Dec-2015, last updated 1 nov 2021
' Purpose   : Returns the directory within the local git repo to which vba modules should be written
' -----------------------------------------------------------------------------------------------------------------------
Function WorkbookToLocalGitFolder(wb As Workbook)
          Dim Folder As String
          Dim FolderGood As Boolean
          Dim Prompt As String
          Dim Res As VbMsgBoxResult
          Dim TempFolder As String
          Dim TempFolderGood As Boolean
          Static UserSaidYes As Boolean
          Dim WorkbookName As String
          Dim XLAFolder As String

1         On Error GoTo ErrHandler

2         XLAFolder = DataFromAuditSheet("LocalReleaseFolder")
3         If Right(XLAFolder, 1) = "\" Then XLAFolder = Left(XLAFolder, Len(XLAFolder) - 1)

4         WorkbookName = wb.Name

5         If Left(LCase(wb.Path), 17) = "c:\solumworkbooks" Or _
              Left(LCase(wb.Path), Len(XLAFolder)) = LCase(XLAFolder) Then
              ' Old way of storing many workbooks in one gigantic repo, https://github.com/PGS62/ExcelVBA
6             Folder = sJoinPath(DataFromAuditSheet("VBAGitFolder"), Left(wb.Name, InStrRev(wb.Name, ".") - 1))
7         ElseIf Left(LCase(wb.Path), 12) = "c:\projects\" And _
              Right(LCase(wb.Path), 10) = "\workbooks" Then
              'New way ion which each workbook that I develop has its own repo https://github.com/PGS62/WorkbookName
              ' and workbooks are in a workbooks folder with vba in vba\workbookname (vba folder same depth as workbooks folder)
8             Folder = Left(wb.Path, Len(wb.Path) - 9) & "vba\" & wb.Name
9         ElseIf Left(LCase(wb.Path), 19) = "c:\projects\cayley\" Then
              'Special case for Cayley project... cope with workbooks in both workbooks and data sub-folders.
10            Folder = Left(wb.Path, 19) & "vba\" & wb.Name
11        Else
12            Throw "cannot determine location of local git folder"
13        End If
          
14        sCreateFolder Folder
15        FolderGood = sFolderIsWritable(Folder)

16        If Not FolderGood Then
17            TempFolder = "c:\temp\VBACode\" + sStringBetweenStrings(WorkbookName, , ".")
18            If sIsErrorString(sCreateFolder(TempFolder)) Then
19                TempFolder = Environ$("TEMP") + "\VBACode\" + sStringBetweenStrings(WorkbookName, , ".")
20                sCreateFolder TempFolder
21            End If
22            TempFolderGood = sFolderIsWritable(TempFolder)
23        End If

24        If TempFolderGood And Not FolderGood Then
25            Prompt = "Folder for export of VBA code for Git source control is set to:" + _
                  vbLf + Folder + vbLf + _
                  "But this folder does not exist or is not writeable. Use a temporary folder instead?" + vbLf + _
                  TempFolder
26            If UserSaidYes Then
27                Res = vbYes
28            Else
29                Res = MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, , , , , , 500)
30            End If
31            If Res = vbYes Then
32                UserSaidYes = True
33                Folder = TempFolder
34            Else
35                Throw "Release aborted"
36            End If
37        End If

38        WorkbookToLocalGitFolder = Folder

39        Exit Function
ErrHandler:
40        Throw "#WorkbookToLocalGitFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub getLatestVersionInfo(ByRef LatestAvailableVersionNumber, ByRef LatestAvailableReleaseDate)

          Dim cn_Version As Variant
          Dim FileName As String
          Dim Headers
          Dim VersionCell As Range

1         On Error GoTo ErrHandler
2         Headers = sArrayTranspose(RangeFromSheet(shAudit, "Headers").Value)
3         cn_Version = sMatch("Version", Headers)
4         If Not IsNumber(cn_Version) Then Throw "Cannot find text ""Version"" in Range ""Headers"" of sheet """ + shAudit.Name + """"
5         Set VersionCell = shAudit.Range("Headers").Cells(2, cn_Version)
6         FileName = DataFromAuditSheet("NetworkReleaseFolder") + "Addins\" & gAddinName & ".xlam"
7         If Not sFileExists(FileName) Then Throw ("Cannot find file '" + FileName + "'")
8         LatestAvailableVersionNumber = sCellContentsFromFileOnDisk(FileName, shAudit.Name, VersionCell.address)
9         LatestAvailableReleaseDate = sFileLastModifiedDate(FileName)

10        Exit Sub
ErrHandler:
11        Throw "#getLatestVersionInfo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckForUpdates
' Author    : Philip Swannell
' Date      : 18-Dec-2015
' Purpose   : Updates to a newer version of SolumAddin
' -----------------------------------------------------------------------------------------------------------------------
Sub CheckForUpdates()
          Dim Installer As String
          Dim LatestAvailableReleaseDate As String
          Dim LatestAvailableVersionNumber As Long
          Dim Prompt As String
          Dim ThisVersion As Long
1         On Error GoTo ErrHandler

2         getLatestVersionInfo LatestAvailableVersionNumber, LatestAvailableReleaseDate

3         ThisVersion = sAddinVersionNumber()
4         If ThisVersion >= LatestAvailableVersionNumber Then
5             Prompt = "You have the latest version of " & gAddinName & ":" + vbLf + CStr(ThisVersion) + " released " + Format$(sAddinReleaseDate(), "dd-mmm-yyyy hh:mm") + "."
6             MsgBoxPlus Prompt, vbOKOnly + vbInformation, "Check for updates", , , , , 320
7             Exit Sub
8         Else
9             Prompt = gAddinName & " " + CStr(LatestAvailableVersionNumber) + " is available (you're using " + CStr(ThisVersion) + ")." + vbLf + vbLf + _
                  "Would you like to quit Excel and install the lastest version of " + gCompanyName + " software, including " & gAddinName & "?"
10            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, "Check for updates", "Yes, upgrade", "No thanks", , , 320) <> vbOK Then Exit Sub
11            Prompt = "Are you sure? You will lose any unsaved work!"
12            If MsgBoxPlus(Prompt, vbOKCancel + vbDefaultButton2 + vbExclamation, , "Go ahead, I've saved my work", "Ah, no do nothing") <> vbOK Then Exit Sub
13            Installer = DataFromAuditSheet("NetworkReleaseFolder") & "Installer\Install.vbs"
14            If Not sFileExists(Installer) Then Throw "Cannot access file '" + Installer + "'"

              Dim AllBooks
              Dim i As Long
              Dim wb As Excel.Workbook
15            For Each wb In Application.Workbooks
16                wb.Saved = True
17            Next
18            AllBooks = WorkbookAndAddInList(3)
19            For i = 1 To sNRows(AllBooks)
20                Application.Workbooks(AllBooks(i, 1)).Saved = True
21            Next i

22            Application.EnableEvents = False
23            For Each wb In Application.Workbooks
24                wb.Close False
25            Next

26            Shell "wscript.exe """ + DataFromAuditSheet("NetworkReleaseFolder") & "Installer\Install.vbs" + """"
27            Application.Quit
28        End If

29        Exit Sub
ErrHandler:
30        Throw "#CheckForUpdates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileMakeLowerCase
' Author    : Philip Swannell
' Date      : 11-Jan-2016
' Purpose   : Creates a new text file, TargetFile, containing the text of SourceFile in lower case.
' Arguments
' SourceFile: Full name (with path) of the source file.
' TargetFile: Full name (with path) of the target file. TargetFile can be the same as SourceFile.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileMakeLowerCase(SourceFile As String, TargetFile As String)
Attribute sFileMakeLowerCase.VB_Description = "Creates a new text file, TargetFile, containing the text of SourceFile in lower case."
Attribute sFileMakeLowerCase.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim FileContents As String
          Dim FSO As Scripting.FileSystemObject
          Dim t1 As TextStream
          Dim t2 As TextStream
1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute SourceFile
3         CheckFileNameIsAbsolute TargetFile
4         Set FSO = New Scripting.FileSystemObject
5         Set t1 = FSO.OpenTextFile(SourceFile, ForReading)
6         FileContents = t1.ReadAll
7         t1.Close
8         Set t2 = FSO.CreateTextFile(TargetFile, True)
9         t2.Write LCase$(FileContents)
10        t2.Close
11        sFileMakeLowerCase = True
12        Set t1 = Nothing: Set t2 = Nothing: Set FSO = Nothing

13        Exit Function
ErrHandler:
14        sFileMakeLowerCase = "#sFileMakeLowerCase (line " & CStr(Erl) + "): " & Err.Description & "!"
15        If Not t1 Is Nothing Then
16            t1.Close
17            Set t1 = Nothing
18        End If
19        If Not t2 Is Nothing Then
20            t2.Close
21            Set t2 = Nothing
22        End If
23        Set FSO = Nothing
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AuditMenuForAddin
' Author    : Philip
' Date      : 24-Dec-2015
' Purpose   : Attached to the Menu... button on the Audit sheet of this workbook.
'             See also method AuditMenu that can be used from the Audit sheet of any workbook
' -----------------------------------------------------------------------------------------------------------------------
Sub AuditMenuForAddin(Optional wb As Excel.Workbook)
1         On Error GoTo ErrHandler

          Dim AddinName As String
          Dim chAVBA As String
          Dim chCheck As String
          Dim chMakeAddin As String
          Dim chRelease As String
          Dim chReleaseNAS As String
          Dim chUnload As String
2         If wb Is Nothing Then Set wb = ActiveWorkbook

3         If ActiveWorkbook Is wb Then
4             AddinName = "me"
5             chCheck = "Check this version against latest release"
6         Else
7             AddinName = sStringBetweenStrings(wb.Name, , ".")
8             chCheck = "Check this version of " + AddinName + " against latest release"
9         End If
          Const chAddRow = "Add &Row to Audit Sheet       (Shift for Comment History)"

10        chMakeAddin = "Make " + AddinName + " an &Addin"
11        chRelease = "Save " + AddinName + " to this &PC"
          Const chVBADiff = "VBA&Diff..."
12        chReleaseNAS = "Release " + AddinName
          Const chReleaseAllSoftware = "Release all " & gCompanyName & " Software"
13        chAVBA = "--Amend VBA code of " + sStringBetweenStrings(wb.Name, , ".") + "..."
          Const chExport = "E&Xport VBA code"
          Const chViewExported = "&View exported VBA code"
14        chUnload = "--&Unload " + AddinName + "..."
          Dim Alternatives
          Dim Chosen
          Dim FaceIDs

15        Alternatives = sArrayStack(chAddRow, chCheck, chMakeAddin, chRelease, chReleaseNAS, chReleaseAllSoftware, chAVBA, chExport, chViewExported, chUnload)
16        FaceIDs = sArrayStack(295, 0, 15023, 3, 0, 0, 293, 13549, 0, 0)
17        Chosen = ShowCommandBarPopup(Alternatives, FaceIDs)
18        Select Case Chosen

              Case Unembellish(chAddRow)
19                AddLineToAuditSheet wb.Worksheets("Audit"), IsShiftKeyDown()
20            Case Unembellish(chCheck)
21                CheckForOverwrite wb, True, True
22            Case Unembellish(chRelease)
23                SaveAddinAndExportVBA wb
24            Case Unembellish(chReleaseAllSoftware)
                  Dim Comment As String
25                GetComment Comment
26                If Comment = "False" Then Exit Sub
27                SaveAddinAndExportVBA wb
28                ReleaseSoftware Comment
29            Case Unembellish(chMakeAddin)
30                CtrlF6Response        'To ensure that focus is on another workbook rather than some application other than Excel! _
                                         (Applies to Office 2013)
31                wb.isAddin = True
32            Case Unembellish(chReleaseNAS)
33                ReleaseToNAS wb, False
34            Case Unembellish(chExport)
35                ExportModules wb, WorkbookToLocalGitFolder(wb), False
36            Case Unembellish(chUnload)
37                UnloadMe
38            Case Unembellish(chViewExported)
39                ViewExportedVBA wb
40            Case Unembellish(chAVBA)
41                If IsInCollection(Application.Workbooks, "AmendVBA.xlam") Then
42                    Application.Run "AmendVBA.XLAM!AmendVBAOfWorkbook", wb
43                Else
44                    Throw "Addin AmendVBA.xlam is not open"
45                End If
46        End Select

47        Exit Sub
ErrHandler:
48        SomethingWentWrong "#AuditMenuForAddin (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ViewExportedVBA
' Author    : Philip Swannell
' Date      : 04-Jan-2016
' Purpose   : Pops Explorer at the appropriate directory - can then use TortoiseGit to check code in - would be good to do this automatically
' -----------------------------------------------------------------------------------------------------------------------
Sub ViewExportedVBA(wb As Excel.Workbook)
1         On Error GoTo ErrHandler
2         Shell "Explorer """ + WorkbookToLocalGitFolder(wb), vbNormalFocus
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ViewExportedVBA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AboutMe
' Author    : Philip Swannell
' Date      : 19-May-2015
' Purpose   : Called from Ribbon to show the user what version they are running.
' -----------------------------------------------------------------------------------------------------------------------
Sub AboutMe()
          Dim Prompt As String
          Dim Res As VbMsgBoxResult
          Dim TimeStamp As String
1         On Error GoTo ErrHandler

2         TimeStamp = Format$(ThrowIfError(sAddinReleaseDate()), "dd-mmm-yyyy")
3         Prompt = "This is " & gAddinName & " version " + Format$(ThrowIfError(sAddinVersionNumber), "#,###") + " dated " + TimeStamp
4         If IsInCollection(Application.Workbooks, gAddinName2 & ".xlam") Then
              Dim LookupTable
              Dim SSUTimeStamp
              Dim SSUVersion
5             LookupTable = sArrayTranspose(Application.Workbooks(gAddinName2 & ".xlam").Worksheets("Audit").Range("Headers").Resize(2).Value2)
6             SSUVersion = Format$(sVLookup("Version", LookupTable), "#,###")
7             SSUTimeStamp = Format$(sVLookup("Date", LookupTable) + sVLookup("Time", LookupTable), "dd-mmm-yyyy")
8             Prompt = Prompt + vbLf + "with " & gAddinName2 & " version " + SSUVersion + " dated " + SSUTimeStamp + "."
9         Else
10            Prompt = Prompt + "."
11        End If

12        Prompt = Prompt + vbLf + vbLf + "Contact: Philip Swannell" + vbLf + "philip.swannell@solum-financial.com" + vbLf + "+44 (0)20 7786 9239"

13        Res = MsgBoxPlus(Prompt, vbInformation + vbYesNoCancel + vbDefaultButton3, "About " + gAddinName, "Check for Updates", gAddinName & " Preferences...", "OK")
14        If Res = vbYes Then
15            CheckForUpdates
16        ElseIf Res = vbNo Then
17            ThisAddinPreferences
18        End If

19        Exit Sub
ErrHandler:
20        SomethingWentWrong "#AboutMe (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnloadMe
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Unload SolumAddin e.g. so that I can edit the ribbon XML
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UnloadMe()
          Const Title = "Unload " & gAddinName
          Dim FileName
          Dim Prompt
1         FileName = ThisWorkbook.FullName
2         Prompt = "Are you sure you want to unload " & gAddinName & "?" + vbLf + "Unsaved changes will be lost!" + vbLf + vbLf + "Last save was " + sDescribeTime(Now() - sFileLastModifiedDate(ThisWorkbook.FullName)) + ", at " + Format$(sFileLastModifiedDate(ThisWorkbook.FullName), "dd-mmm-yy hh:mm:ss") + "."
3         If MsgBoxPlus(Prompt, vbExclamation + vbOKCancel + vbDefaultButton2, Title, "Unload") <> vbOK Then Exit Sub
4         Application.Addins(gAddinName).Installed = False
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDescribeTime
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Friendly description of a time interval
' -----------------------------------------------------------------------------------------------------------------------
Function sDescribeTime(NumDays As Double)
Attribute sDescribeTime.VB_Description = "An approximate description of a point in time NumDays in the past. E.g. ""About 3 months ago"" or ""A few seconds ago""."
Attribute sDescribeTime.VB_ProcData.VB_Invoke_Func = " \n28"
          Const OneSecond = 1 / 24 / 60 / 60
          Const OneMinute = 1 / 24 / 60
          Const OneHour = 1 / 24

          Dim nDays As Long
          Dim nHours As Long
          Dim nMinutes As Long
          Dim nMonths As Long
          Dim nYears As Long

1         On Error GoTo ErrHandler
2         nYears = WorksheetFunction.RoundDown(NumDays / 365, 0)
3         nMonths = WorksheetFunction.RoundDown(NumDays / 365 * 12 - nYears * 12, 0)
4         nDays = WorksheetFunction.RoundDown(NumDays - nYears * 365 - nMonths * 365 / 12, 0)
5         nHours = WorksheetFunction.RoundDown(NumDays / OneHour - nYears * 365 * 24 - nMonths * 365 / 12 * 24 - nDays * 24, 0)
6         nMinutes = WorksheetFunction.RoundDown(NumDays / OneMinute - nYears * 365 * 24 * 60 - nMonths * 365 / 12 * 24 * 60 - nDays * 24 * 60 - nHours * 60, 0)

7         If NumDays <= 0 Then
8             Throw "NumDays must be positive"
9         ElseIf NumDays < 30 * OneSecond Then
10            sDescribeTime = "a few seconds ago"
11        ElseIf NumDays < 90 * OneSecond Then
12            sDescribeTime = "about a minute ago"
13        ElseIf NumDays < OneHour Then
14            sDescribeTime = "about " + CStr(CLng(NumDays / OneMinute)) + " minutes ago"
15        ElseIf NumDays < OneHour * 5 Then
16            sDescribeTime = "about " + CStr(nHours) + " hour" + IIf(nHours > 1, "s", vbNullString) + IIf(nMinutes <> 0, " " + CStr(nMinutes) + " minute" + IIf(nMinutes > 1, "s", vbNullString), vbNullString) + " ago"
17        ElseIf NumDays < 1 Then
18            sDescribeTime = "about " + CStr(nHours) + " hours ago"
19        ElseIf NumDays < 3 Then
20            sDescribeTime = "about " + CStr(nDays) + " day" + IIf(nDays > 1, "s", vbNullString) + IIf(nHours <> 0, " " + CStr(nHours) + " hour" + IIf(nHours > 1, "s", vbNullString), vbNullString) + " ago"
21        ElseIf NumDays <= 31 Then
22            sDescribeTime = "about " + CStr(nDays) + " day" + IIf(nDays > 1, "s", vbNullString) + " ago"
23        ElseIf NumDays < 180 Then
24            sDescribeTime = "about " + CStr(nMonths) + " month" + IIf(nMonths > 1, "s", vbNullString) + IIf(nDays <> 0, " " + CStr(nDays) + " day" + IIf(nDays > 1, "s", vbNullString), vbNullString) + " ago"
25        ElseIf NumDays < 365 Then
26            sDescribeTime = "about " + CStr(nMonths) + " months ago"
27        ElseIf NumDays < 5 * 365 Then
28            sDescribeTime = "about " + CStr(nYears) + " year" + IIf(nYears > 1, "s", vbNullString) + IIf(nMonths <> 0, " " + CStr(nMonths) + " month" + IIf(nMonths > 1, "s", vbNullString), vbNullString) + " ago"
29        Else
30            sDescribeTime = "about " + CStr(nYears) + " years ago"
31        End If
32        Exit Function
ErrHandler:
33        sDescribeTime = "#sDescribeTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RegisterButton_Click
' Author    : Philip Swannell
' Date      : 21-May-2015
' Purpose   : Register functions from button on Help sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterButton_Click()
          Dim Warning As String
1         On Error GoTo ErrHandler

2         If Not sArraysIdentical(shHelp.Range("TheData").Columns(1).Value, sSortedArray(shHelp.Range("TheData").Columns(1))) Then
3             Throw "Range 'TheData' is not correctly sorted on FunctionName. Please fix before releasing."
4         End If

5         RefreshRibbon
6         RefreshIntellisenseSheet shHelp.Range("TheData"), shIntellisense
7         UninstallIntellisense
8         InstallIntellisense

9         RegisterFunctionsWithFunctionWizard Warning
10        If Warning <> "OK" Then
11            MsgBoxPlus Warning, vbExclamation, "R1egister " & gCompanyName & " Functions"
12        Else
13            MsgBoxPlus "All functions registered correctly." + vbLf + vbLf + "Sheet '_IntelliSense_' was updated correctly", vbInformation, "Register " & gCompanyName & " Functions"
14        End If
15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#RegisterButton_Click (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RegisterFunctionsWithFunctionWizard
' Author    : Philip Swannell
' Date      : 03-May-2015
' Purpose   : Uses the data held on the "Help" sheet to make the s... functions appear
'             in the Excel Function Wizard. Returns an error string is one or more functions
'             fail to register. We ignore that error in the Workbook_Open but don't ignore it
'             in method SaveAddinAndExportVBA.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterFunctionsWithFunctionWizard(ByRef Warning As String)
          Dim ArgumentDescriptions() As String
          Dim Category As String
          Dim Col_ArgumentDescription As Variant
          Dim Col_Category As Variant
          Dim Col_Description As Variant
          Dim Col_FnName As Variant
          Dim Col_NumArgs As Variant
          Dim ColumnHeaders As Variant
          Dim Description As String
          Dim FailedFunctions As String
          Dim FunctionName As String
          Dim HelpData As Variant
          Dim i As Long
          Dim j As Long
          Dim NumArgs As Long
          Dim NumFailures As Long

1         On Error GoTo ErrHandler

2         ColumnHeaders = Application.WorksheetFunction.Transpose(shHelp.Range("TheData").Rows(0).Value)
3         Col_FnName = ThrowIfError(sMatch("FunctionName", ColumnHeaders))
4         Col_NumArgs = ThrowIfError(sMatch("N", ColumnHeaders))
5         Col_Description = ThrowIfError(sMatch("FunctionWizard Description", ColumnHeaders))
6         Col_ArgumentDescription = ThrowIfError(sMatch("Argument 1", ColumnHeaders))
7         Col_Category = ThrowIfError(sMatch("TheCategories", ColumnHeaders))
8         HelpData = GetHelpData()

9         For i = 1 To sNRows(HelpData)
10            If Not IsEmpty(HelpData(i, Col_Description)) Then
11                FunctionName = CStr(HelpData(i, Col_FnName))
12                Description = CStr(HelpData(i, Col_Description))
13                Category = gCompanyName & " " + CStr(HelpData(i, Col_Category))
14                If IsEmpty(HelpData(i, Col_ArgumentDescription)) Then

15                    If Not SafeMacroOptions(FunctionName, Description, Category) Then
16                        NumFailures = NumFailures + 1
17                        FailedFunctions = FailedFunctions + vbLf + FunctionName
18                    End If

19                Else
20                    NumArgs = HelpData(i, Col_NumArgs)
21                    ReDim ArgumentDescriptions(1 To NumArgs)
22                    For j = 1 To NumArgs
23                        ArgumentDescriptions(j) = CStr(HelpData(i, Col_ArgumentDescription - 1 + j))
24                    Next j
25                    If Not SafeMacroOptions(FunctionName, Description, Category, ArgumentDescriptions) Then
26                        NumFailures = NumFailures + 1
27                        FailedFunctions = FailedFunctions + vbLf + FunctionName
28                    End If
29                End If
30            End If
31        Next i
32        If NumFailures = 0 Then
33            Warning = "OK"
34        Else
35            Warning = CStr(NumFailures) + _
                  " function" + IIf(NumFailures > 1, "s", vbNullString) + " failed to register." + vbLf + FailedFunctions + _
                  vbLf + vbLf + "See VBA immediate window for more details!"
36        End If

37        Exit Sub
ErrHandler:
38        Throw "#RegisterFunctionsWithFunctionWizard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeMacroOptions
' Author    : Philip Swannell
' Date      : 03-May-2015
' Purpose   : Error handling around Application.MacroOptions
' -----------------------------------------------------------------------------------------------------------------------
Private Function SafeMacroOptions(FunctionName As String, Description As String, Category As String, Optional ArgumentDescriptions As Variant)
          Dim ErrorString As String
1         On Error GoTo ErrHandler
2         Application.MacroOptions FunctionName, Description, , , , , Category, , , , ArgumentDescriptions
3         SafeMacroOptions = True
4         Exit Function
ErrHandler:
5         ErrorString = "Error in Application.MacroOptions for function " + FunctionName + ": " + Err.Description
6         Debug.Print ErrorString
7         SafeMacroOptions = False
End Function

Private Function Is64bit() As Boolean
#If Win64 Then
1         Is64bit = True
#End If
End Function

Sub UninstallIntellisense()
          Dim a As AddIn
          Dim AddinName As String
1         On Error GoTo ErrHandler
2         If Is64bit() Then
3             AddinName = "ExcelDna.IntelliSense64.xll"
4         Else
5             AddinName = "ExcelDna.IntelliSense.xll"
6         End If

7         For Each a In Application.Addins
8             If LCase$(a.Name) = LCase$(AddinName) Then
9                 If a.Installed Then a.Installed = False
10                Exit Sub
11            End If
12        Next

13        Exit Sub
ErrHandler:
14        Throw "#UninstallIntellisense (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InstallIntellisense
' Author     : Philip Swannell
' Date       : 21-Jan-2019
' Purpose    : Installs the Intellisense addin, assumed to be available on the C drive \ProgramData\<gCompanyName>\Addins\
'              See https://github.com/Excel-DNA/IntelliSense
'              No need to call this method in the Workbook_Open code sinvce the Install.vbs installer does the installation
'              (by editing the Registry) but call from RegisterButton_Click - useful when editing the help for functions.
' -----------------------------------------------------------------------------------------------------------------------
Sub InstallIntellisense()
          Dim a As AddIn
          Dim AddinFullName As String
          Dim AddinName As String
          Dim Folder As String
          Dim j As Long

1         On Error GoTo ErrHandler
2         Folder = "C:\ProgramData\" & gCompanyName & "\ExcelDNA\"

3         If Is64bit() Then
4             AddinName = "ExcelDna.IntelliSense64.xll"
5         Else
6             AddinName = "ExcelDna.IntelliSense.xll" 'CHECK THIS!
7         End If
8         AddinFullName = Folder + AddinName

9         For j = 1 To 2
10            For Each a In Application.Addins
11                If LCase$(a.FullName) = LCase$(AddinFullName) Then
12                    If Not a.Installed Then a.Installed = True
13                    Exit Sub 'Already installed so exit
14                ElseIf LCase$(a.Name) = LCase$(AddinName) Then
15                    a.Installed = False 'installed from a different location, so uninstall it
16                End If
17            Next
18            If sFileExists(AddinFullName) Then
19                Application.Addins.Add AddinFullName
20                Exit Sub
21            End If
22        Next j

23        Exit Sub
ErrHandler:
24        Debug.Print "#InstallIntellisense (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RefreshIntellisenseSheet
' Author     : Philip Swannell
' Date       : 21-Jan-2019
' Purpose    : Writes from scratch the _IntelliSense_ worksheet that's used by the addin ExcelDna.IntelliSense.xll
'              (or ExcelDna.IntelliSense64.xll). Source of the help data is the worksheet Help. For the time being we are
'              thus storing all that data twice. Storing only once would involve re-writing all methods that currently
'              read the Help sheet: CleanMe, GetHelpData, HelpVBE, InsertFunctionAtActiveCell,
'              RegisterFunctionsWithFunctionWizard, ShowHelpForFunction, TheRibbon_getContent, TheRibbon_getScreentip,
'              TheRibbon_getSupertip, TheRibbon_onAction
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshIntellisenseSheet(SourceRange As Range, TargetSheet As Worksheet)
          Dim cn_ArgNames As Long
          Dim cn_Category As Long
          Dim cn_FirstArg As Long
          Dim cn_Fn As Long
          Dim cn_FnDescription As Long
          Dim cn_FnExtraDescription As Long
          Dim cn_NumArgs As Long
          Dim Fn As String
          Dim FnArgs As Variant
          Dim FnDescription As String
          Dim i As Long
          Dim j As Long
          Dim NumArgs As Long
          Dim Titles
          Dim WriteRow As Long

1         On Error GoTo ErrHandler
2         Titles = sArrayTranspose(SourceRange.Rows(0))
3         cn_FnDescription = ThrowIfError(sMatch("FunctionWizard Description", Titles))
4         cn_FnExtraDescription = ThrowIfError(sMatch("Extra Description", Titles))
5         cn_FirstArg = ThrowIfError(sMatch("Argument 1", Titles))
6         cn_ArgNames = ThrowIfError(sMatch("TheExplanationTitles", Titles))
7         cn_Fn = ThrowIfError(sMatch("FunctionName", Titles))
8         cn_Category = ThrowIfError(sMatch("TheCategories", Titles))
9         cn_NumArgs = ThrowIfError(sMatch("N", Titles))

10        With TargetSheet
11            .UsedRange.EntireColumn.Delete
12            .Cells(1, 1).Value = "FunctionInfo"
13            .Cells(1, 2).Value = "'1.0"
14            WriteRow = 2

15            For i = 1 To SourceRange.Rows.Count
16                If SourceRange.Cells(i, cn_Category) <> "Keyboard Shortcuts" Then
17                    Fn = SourceRange.Cells(i, cn_Fn).Value
18                    FnDescription = SourceRange.Cells(i, cn_FnDescription).Value
19                    If Len(SourceRange.Cells(i, cn_FnExtraDescription)) > 0 Then
20                        FnDescription = FnDescription + " More details at " + gCompanyName + " > Browse Functions."
21                    End If
22                    FnDescription = Replace(FnDescription, vbLf, " ") 'because the text in the intellisense popups does not respect line breaks, tried html <br> but no joy.

23                    FnArgs = SourceRange.Cells(i, cn_ArgNames)
24                    FnArgs = Replace(FnArgs, ",...", vbNullString)
25                    FnArgs = sStringBetweenStrings(FnArgs, "(", ")")
26                    If FnArgs = vbNullString Then
27                        NumArgs = 0
28                    Else
29                        FnArgs = sTokeniseString(CStr(FnArgs))
30                        NumArgs = sNRows(FnArgs)
31                    End If
32                    If NumArgs <> SourceRange.Cells(i, cn_NumArgs) Then Throw "Detected mismatch in argument count for function " + Fn

33                    .Cells(WriteRow, 1).Value = Fn
34                    .Cells(WriteRow, 2).Value = FnDescription
                          
35                    For j = 1 To NumArgs
36                        .Cells(WriteRow, 4 + 2 * (j - 1)).Value = FnArgs(j, 1)
37                        .Cells(WriteRow, 5 + 2 * (j - 1)).Value = Replace(SourceRange(i, cn_FirstArg + j - 1).Value, vbLf, " ")
38                    Next j
39                    WriteRow = WriteRow + 1

40                End If
41            Next i

42            With .UsedRange
43                .ColumnWidth = 40
44                .HorizontalAlignment = xlHAlignLeft
45                .VerticalAlignment = xlVAlignCenter
46                .WrapText = True
47                .RowHeight = 36.6
48                AddGreyBorders .Offset(0)
49            End With
50        End With

51        Exit Sub
ErrHandler:
52        Throw "#RefreshIntellisenseSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
