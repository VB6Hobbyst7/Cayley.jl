Attribute VB_Name = "modSCRiPTUtils"
Option Explicit
'Public Const gRSourcePath = "c:/ProgramData/" & gCompanyName & "/RSource/"      'in Unix notation, with trailing /
Public Const gBasePath = "c:\ProgramData\" & gCompanyName & "\"  'In Windows notation, with trailing \. CHANGING THIS? Also change folder data held on the Audit sheet.

'CHANGING CONSTANT gPackages?
'Then also change default value for first argument to R method InstallPackages!!!
Public Const gPackages = "BB,data.table,digest,doParallel,foreach,Matrix,nleqslv,pcaPP,plyr,randtoolbox,reshape,rngWELL,jsonlite,iterators,itertools"    'packages required by either SCRiPTMain.R or SolumAddin.R
Public Const gPackagesSAI = "pcaPP,data.table,reshape,stats,iterators,itertools"    'Packages required by code in SolumAddin.R
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modSCRiPTUtils
' Author    : Philip Swannell
' Date      : 15-June-2016
' Purpose   : Common code to be called from both the SCRIPT and SCRiPT_MarketData and Cayley
'             workbooks
'             Much of the functionality of this module moved to SCRiPTUtils.xlam 13 Dec 2016
' -----------------------------------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : gRSourcePath
' Author     : Philip Swannell
' Date       : 12-Dec-2018
' Purpose    : Get the location of the RSource directory from the SCRiPT workbook if its open, if not check for another
'              workbook which has the correct sheet name and range name, if not look in the Registry in the location where the SCRiPT workbook saves to, finally use a default value...
' -----------------------------------------------------------------------------------------------------------------------
Function gRSourcePath() As String
          Dim Res As String
          Dim wb As Excel.Workbook
          Dim BookName As String

1         On Error Resume Next

2         BookName = "SCRiPT.xlsm"
3         Res = Application.Workbooks(BookName).Worksheets("Config").Range("RSourcePath").Value
4         On Error GoTo ErrHandler
5         If Res = vbNullString Then
6             For Each wb In Application.Workbooks
7                 If IsInCollection(wb.Worksheets, "Config") Then
8                     If IsInCollection(wb.Worksheets("Config").Names, "RSourcePath") Then
9                         BookName = wb.Name
10                        Res = wb.Worksheets("Config").Range("RSourcePath").Value
11                        Exit For
12                    End If
13                End If
14            Next
15        End If
16        If Res <> vbNullString Then
17            If Not sFolderExists(Res) Then Throw "Workbook " + BookName + " sheet Config range RSourcePath is set to '" + Res + "' but that folder does not exist"
18            If Not sFileExists(sJoinPath(Res, "SCRiPTMain.R")) Then Throw "Workbook " + BookName + " sheet Config range RSourcePath is set to '" + Res + "' but that folder does not contain the file 'SCRiPTMain.R'"
19        End If

20        If Res = vbNullString Then
21            Res = GetSetting(gCompanyName & "Config", "SCRiPT", "RSourcePath", vbNullString)
22            If Res <> vbNullString Then
23                Res = sParseArrayString(Res)(1, 1)
                  Dim Reglocation As String
24                Reglocation = "Registry key 'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\" + gCompanyName & "Config\SCRiPT\RSourcePath"
25                If Not sFolderExists(Res) Then Throw "Registry location '" + Reglocation + "' defines a folder '" + Res + "' but that folder does not exist"
26                If Not sFileExists(sJoinPath(Res, "SCRiPTMain.R")) Then Throw "Registry location '" + Reglocation + "' defines a folder '" + Res + "' but that folder does not contain the file 'SCRiPTMain.R'"
27            End If
28        End If

29        If Res = vbNullString Then
30            Res = "c:\ProgramData\" & gCompanyName & "\RSource"
31            If Not sFolderExists(Res) Then Throw "Cannot find folder '" + Res + "'"
32            If Not sFileExists(sJoinPath(Res, "SCRiPTMain.R")) Then Throw "Folder '" + Res + "' does not contain the file 'SCRiPTMain.R'"
33        End If
34        Res = Replace(Res, "\", "/")
35        If Right$(Res, 1) <> "/" Then Res = Res + "/"
36        gRSourcePath = Res
37        Exit Function
ErrHandler:
38        Throw "#gRSourcePath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunReleaseCleanup
' Author    : Philip Swannell
' Date      : 11-Jul-2017
' Purpose   : Call a workbook's release clean up or do a "default" cleanup.
' -----------------------------------------------------------------------------------------------------------------------
Sub RunReleaseCleanup(wb As Excel.Workbook)
          Dim i As Long
1         On Error GoTo ErrHandler

2         If Application.DisplayFormulaBar Then Application.FormulaBarHeight = 1
3         AddAuditSheetToBook wb

4         If True Then
5             ThrowIfError Application.Run("'" + Replace(wb.Name, "'", "''") & "'!ReleaseCleanup")
6             Exit Sub
7         Else
ResumeHere:
              Dim ws As Worksheet
8             For Each ws In wb.Worksheets
9                 If ws.Visible = xlSheetVisible Then
10                    Application.Goto ws.Cells(1, 1)
11                    ActiveWindow.DisplayGridlines = False
12                    ActiveWindow.DisplayHeadings = False
13                End If
14                ws.Protect , True, True
15            Next
              'Make the leftmost tab active
16            For i = 1 To wb.Worksheets.Count
17                If wb.Worksheets(i).Visible = xlSheetVisible Then
18                    Application.Goto wb.Worksheets(i).Cells(1, 1)
19                    Exit For
20                End If
21            Next i
22        End If
23        Exit Sub
ErrHandler:
24        If Err.Number = 1004 Then GoTo ResumeHere
25        Throw "#RunReleaseCleanup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FixLinksForActiveBook
' Author     : Philip Swannell
' Date       : 11-Dec-2017
' Purpose    : Change out-of-date links to SolumAddin.xlam and SolumSCRiPTUtils.xlam for the active workbook
' -----------------------------------------------------------------------------------------------------------------------
Sub FixLinksForActiveBook()
          Dim LinksChanged As Boolean
          Dim Prompt
          Dim ReferencesChanged As Boolean

          Const Title = "Fix Links (" + gAddinName + ")"

1         On Error GoTo ErrHandler
2         If Not ActiveWorkbook Is Nothing Then
3             Prompt = "For the Active workbook, fix links and VBA references to " & gAddinName & " and " & gAddinName2 & "?"
4             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title, "Yes, fix", "No, do nothing") <> vbOK Then Exit Sub
5             FixLinks ActiveWorkbook, LinksChanged, ReferencesChanged
6             If LinksChanged And ReferencesChanged Then
7                 Prompt = "All done. Links and VBA references were changed."
8             ElseIf LinksChanged Then
9                 Prompt = "All done. Links were changed, but no VBA references needed to change."
10            ElseIf ReferencesChanged Then
11                Prompt = "All done. VBA references were changed, but no links needed to change."
12            Else
13                Prompt = "All done. Neither links nor VBA references needed to change."
14            End If
15            MsgBoxPlus Prompt, vbOKOnly + vbInformation, Title
16        End If
17        Exit Sub
ErrHandler:
18        Throw "#FixLinksForActiveBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FixLinks
' Author     : Philip Swannell
' Date       : 11-Dec-2017
' Purpose    : Change the Excel links and VBA references in a workbook for our two addins moving from their old location
'              (or any incorrect location) to their new location.
' Parameters :
'  wb               : The workbook
'  LinksChanged     : By reference to tell the caller if links had to be changed
'  ReferencesChanged: By reference to tell the caller if references had to be changed
'  SilentMode       : Pass as True when calling from clsApp. Function then does not throw errors (but returns them as strings) and if
'                     Access is not trusted to the VBProject, function does not attempt to fix bad references
' -----------------------------------------------------------------------------------------------------------------------
Sub FixLinks(wb As Excel.Workbook, Optional ByRef LinksChanged As Boolean, Optional ByRef ReferencesChanged As Boolean, Optional SilentMode As Boolean)

          Dim i As Long
          Dim LinkSources
          Dim SPHArray() As clsSheetProtectionHandler
          Dim ThisLinkSource As String
          Const Path1 = "c:\ProgramData\" & gCompanyName & "\Addins\" & gAddinName & ".xlam"
          Const Path2 = "c:\ProgramData\" & gCompanyName & "\Addins\" & gAddinName2 & ".xlam"
          Dim CopyOfErr As String
          Dim OrigDir As String
          Dim TrustError As String

1         On Error GoTo ErrHandler
2         OrigDir = CurDir$()

3         ReDim SPHArray(1 To wb.Worksheets.Count)

4         TrustError = "To use this macro, access to the VBA project object model must be allowed." + vbLf + _
              "Please allow it using the Excel menus: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access to the VBA project object model"

          'In SilentMode we fix links even if we can't fix references
5         If Not SilentMode Then If Not AccessTrusted() Then Throw TrustError, True
6         LinkSources = wb.LinkSources

7         If Not LCase$(ThisWorkbook.FullName) = LCase$(Path1) Then Throw gAddinName & " is installed in the wrong location." + vbLf + _
              "It is installed at:" + vbLf + ThisWorkbook.FullName + vbLf + "It should be installed at:" + vbLf + Path1 + vbLf + vbLf + "Please re-install the " & gCompanyName & " software.", True
8         If Not sFileExists(Path2) Then Throw "Cannot find file: " + vbLf + Path2 + vbLf + vbLf + "Please re-install the " & gCompanyName & " software.", True

9         If Not IsEmpty(LinkSources) Then
10            For i = 1 To wb.Worksheets.Count
11                Set SPHArray(i) = CreateSheetProtectionHandler(wb.Worksheets(i))
12            Next

13            LinkSources = sArrayTranspose(LinkSources)    'makes 2-dimensional and 1-based
14            For i = 1 To sNRows(LinkSources)
15                ThisLinkSource = LinkSources(i, 1)
16                If LCase$(sSplitPath(ThisLinkSource)) = sSplitPath(LCase$(Path1)) Then
17                    If LCase$(ThisLinkSource) <> LCase$(Path1) Then
18                        ChDir sSplitPath(Path1, False)
19                        wb.ChangeLink Name:=ThisLinkSource, NewName:=sSplitPath(Path1), Type:=xlExcelLinks
20                        LinksChanged = True
21                    End If
22                End If
23                If LCase$(sSplitPath(ThisLinkSource)) = sSplitPath(LCase$(Path2)) Then
24                    If LCase$(ThisLinkSource) <> LCase$(Path2) Then
25                        ChDir sSplitPath(Path2, False)
26                        wb.ChangeLink Name:=ThisLinkSource, NewName:=sSplitPath(Path2), Type:=xlExcelLinks
27                        LinksChanged = True
28                    End If
29                End If
30            Next i
31        End If

32        If CurDir$() <> OrigDir Then ChDir OrigDir

          Dim Do1 As Boolean
          Dim Do2 As Boolean
          Dim R As Reference
          Dim r1 As Reference
          Dim r2 As Reference

33        If SilentMode Then If Not AccessTrusted() Then Throw TrustError, False

34        For Each R In wb.VBProject.References
35            If R.IsBroken Then
36                If LCase$(sSplitPath(R.FullPath)) = LCase$(sSplitPath(Path1)) Then
37                    Do1 = True
38                    Set r1 = R
39                End If
40                If LCase$(sSplitPath(R.FullPath)) = LCase$(sSplitPath(Path2)) Then
41                    Do2 = True
42                    Set r2 = R
43                End If
44            End If
45        Next

46        If Do1 Then
47            ReferencesChanged = True
48            wb.VBProject.References.Remove r1
49            wb.VBProject.References.AddFromFile Path1
50        End If

51        If Do2 Then
52            ReferencesChanged = True
53            wb.VBProject.References.Remove r2
54            wb.VBProject.References.AddFromFile Path2
55        End If

56        Exit Sub
ErrHandler:
57        CopyOfErr = "#FixLinks (line " & CStr(Erl) + "): " & Err.Description & "!"
58        If OrigDir <> vbNullString Then If CurDir$() <> OrigDir Then ChDir OrigDir
59        If Not SilentMode Then
60            Throw CopyOfErr
61        End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AccessTrusted
' Author     : Philip Swannell
' Date       : 12-Dec-2017
' Purpose    : Test if user settings allow manipulation of VBProjects
' -----------------------------------------------------------------------------------------------------------------------
Private Function AccessTrusted() As Boolean
1         On Error GoTo ErrHandler

          Dim vbp As VBProject
2         Set vbp = ThisWorkbook.VBProject
3         AccessTrusted = True
4         Exit Function
ErrHandler:
5         AccessTrusted = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : am
' Author     : Philip Swannell
' Date       : 09-Sep-2021
' Purpose    : Convenience when developing workbooks which should have no links to this addin.
' -----------------------------------------------------------------------------------------------------------------------
Sub am()
1         AuditMenu
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AuditMenu
' Author    : Philip Swannell
' Date      : 16-Jun-2016
' Purpose   : Assign to "Menu..." button on the Audit sheet of workbooks for which we want
'             to have decent version-control
' -----------------------------------------------------------------------------------------------------------------------
Sub AuditMenu()
          Const chAddRow = "Add &Row to Audit Sheet       (Shift for Comment History)"
          Const chCheck = "Check this version against latest release"
          Const chCleanup = "&Prepare this workbook for release"
          Const chUnprotect = "&Unprotect all sheets of the workbook"
          Const chExport = "E&xport VBA code of this workbook"
          Const chViewExported = "&View exported VBA code"
          Const chRelease = "Release this &workbook"
          Const chAVBA = "Amend VBA code..."
          Const chOldStyleArrays = "Search this workbook for old-style array formulas"
          Const chMakeAddin = "Make me an AddIn"
          Dim TheChoices
          Dim FaceIDs
          Dim Chosen
          Dim isAddin As Boolean

1         On Error GoTo ErrHandler
2         shAudit.Calculate        'ensure the cell that shows the FullName of the workbook is recalculated
3         TheChoices = sArrayStack(chAddRow, chCheck, chCleanup, chRelease, chUnprotect, "--" + chExport, chViewExported, chAVBA, chOldStyleArrays)
4         FaceIDs = sArrayStack(295, 0, 108, 22, 162, 13549, 0, 293, 9362)
5         isAddin = LCase(Right(ActiveWorkbook.Name, 5)) = ".xlam"
6         If isAddin Then
7             TheChoices = sArrayStack(TheChoices, "--" & chMakeAddin)
8             FaceIDs = sArrayStack(FaceIDs, 15023)
9         End If

10        Chosen = ShowCommandBarPopup(TheChoices, FaceIDs)

11        Select Case Chosen
              Case Unembellish(chAddRow)
12                AddLineToAuditSheet ActiveSheet, IsShiftKeyDown()
13            Case Unembellish(chCheck)
14                CheckForOverwrite ActiveWorkbook, False, True
15            Case Unembellish(chCleanup)
16                RunReleaseCleanup ActiveWorkbook
17            Case Unembellish(chRelease)
18                If ActiveWorkbook.VBProject.Protection = 1 Then
19                    MsgBoxPlus "Please unlock the VBA code in the VBEditor (Alt F11)", vbOKOnly
                      'Activate the VBIDE and close windows
20                    Application.SendKeys "%{F11}%W{UP}{RETURN}"
21                    Exit Sub
22                End If
23                RunReleaseCleanup ActiveWorkbook
24                ReleaseWorkbook ActiveWorkbook
25            Case Unembellish(chUnprotect)
26                UnprotectSheetsOfBook ActiveWorkbook
27            Case Unembellish(chExport)
28                ExportModules ActiveWorkbook, WorkbookToLocalGitFolder(ActiveWorkbook), False
29            Case Unembellish(chViewExported)
30                ViewExportedVBA ActiveWorkbook
31            Case Unembellish(chAVBA)
32                If IsInCollection(Application.Workbooks, "AmendVBA.xlam") Then
33                    Application.Run "AmendVBA.XLAM!AmendVBAOfWorkbook", ActiveWorkbook
34                Else
35                    Throw "Addin AmendVBA.xlam is not open"
36                End If
37            Case Unembellish(chOldStyleArrays)
38                SearchWorkbookFormulas , , ActiveWorkbook, True
39            Case Unembellish(chMakeAddin)
40                ActiveWorkbook.isAddin = True
41        End Select

42        Exit Sub
ErrHandler:
43        SomethingWentWrong "#AuditMenu (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnprotectSheetsOfBook
' Author    : Philip Swannell
' Date      : 08-Jul-2016
' Purpose   : Unprotect all sheets of a workbook and expand any grouping buttons on the sheet.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UnprotectSheetsOfBook(wb As Excel.Workbook)
          Dim oldVisState
          Dim origSheet As Worksheet
          Dim SUH As clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set origSheet = ActiveSheet
4         For Each ws In wb.Worksheets
5             ws.Unprotect
6             oldVisState = ws.Visible
7             ws.Visible = xlSheetVisible
8             ws.Activate
9             GroupingButtonDoAllOnSheet ws, True
10            ActiveWindow.DisplayHeadings = True
11            ws.Visible = oldVisState
12        Next
13        origSheet.Activate
14        Exit Sub
ErrHandler:
15        Throw "#UnprotectSheetsOfBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRightColWhereLeftColIsWeekDay
' Author    : Philip Swannell
' Date      : 17-Jun-2016
' Purpose   : Useful for processing calls to the Bloomberg function BDH for which the
'             argument "Days=WeekDays" does not always do what it says on the tin, notably
'             when requesting data for mid-East currencies
' -----------------------------------------------------------------------------------------------------------------------
Function sRightColWhereLeftColIsWeekDay(ByVal TheData)
1         On Error GoTo ErrHandler

2         Force2DArrayR TheData
          Dim AnyBad As Boolean
          Dim ChooseVector() As Boolean
          Dim i As Long

3         If sNCols(TheData) < 2 Then
4             sRightColWhereLeftColIsWeekDay = TheData
5             Exit Function
6         End If

7         ReDim ChooseVector(1 To sNRows(TheData), 1 To 1)
8         For i = 1 To sNRows(TheData)
9             If IsNumberOrDate(TheData(i, 1)) Then
10                If TheData(i, 1) Mod 7 > 1 Then
11                    ChooseVector(i, 1) = True
12                End If
13            End If
14            If Not ChooseVector(i, 1) Then AnyBad = True
15        Next i
16        If AnyBad Then
17            sRightColWhereLeftColIsWeekDay = sMChoose(sSubArray(TheData, 1, 2, , 1), ChooseVector)
18        Else
19            sRightColWhereLeftColIsWeekDay = sSubArray(TheData, 1, 2, , 1)
20        End If

21        Exit Function
ErrHandler:
22        sRightColWhereLeftColIsWeekDay = "#sRightColWhereLeftColIsWeekDay (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetCalculationToManual
' Author    : Philip Swannell
' Date      : 27-Oct-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Sub SetCalculationToManual()
          Dim Prompt
1         On Error GoTo ErrHandler
2         If Application.Calculation <> xlCalculationManual Then
3             Prompt = "Workbook Calculation (set via File > Options > Formulas > Calculation options) must be 'Manual'. Set to Manual now?"
4             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, gAddinName, "Yes, set to Manual", "No, leave as is", , , , , , 20, vbOK) = vbOK Then
5                 Application.Calculation = xlCalculationManual
6             Else
7                 Throw "Workbook Calculation must be Manual"
8             End If
9         End If
10        Exit Sub
ErrHandler:
11        Throw "#SetCalculationToManual (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveConfigToRegistry
' Author    : Philip Swannell
' Date      : 06-Nov-2016
' Purpose   : Persist a two-column section of a worksheet to the Registry... Useful for Config sheets
' -----------------------------------------------------------------------------------------------------------------------
Sub SaveConfigToRegistry(R As Range, BookName As String)
          Dim a As Range
          Dim c As Range
          Dim NorL As String
1         On Error GoTo ErrHandler
2         For Each a In R.Areas
3             If a.Columns.Count <> 2 Then Throw "R must be a range with two columns"
4         Next a
5         On Error Resume Next
          'Delete what's in the registry at this key since we might have changed the cell names or labels
6         DeleteSetting gCompanyName & "Config", BookName
7         On Error GoTo ErrHandler
8         For Each a In R.Areas
9             For Each c In a.Columns(2).Cells
                  'Use MakeArrayString as it provides a way of preserving type information, e.g. distinguish between Empty and Null string, text TRUE and Boolean TRUE etc.
10                NorL = NameOrLabel(c)
11                If NorL <> vbNullString Then
12                    SaveSetting gCompanyName & "Config", BookName, NorL, CStr(sMakeArrayString(c.Value2))
13                End If
14            Next c
15        Next a
16        Exit Sub
ErrHandler:
17        Throw "#SaveConfigToRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetConfigFromRegistry
' Author    : Philip Swannell
' Date      : 06-Nov-2016
' Purpose   : Read a two-column range back from the Registry, Note that where cells contain formulas we don't write back...
'             Called from SCRiPT, Cayley, SCRiPTMarketData so don't make Private.
' -----------------------------------------------------------------------------------------------------------------------
Sub GetConfigFromRegistry(R As Range, BookName As String)
          Dim a As Range
          Dim c As Range
          Dim Res As String
          Dim SPH As clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         For Each a In R.Areas
3             If a.Columns.Count <> 2 Then Throw "R must be a range with two columns"
4         Next a

5         Set SPH = CreateSheetProtectionHandler(R.Parent)

6         For Each a In R.Areas
7             For Each c In a.Columns(2).Cells
8                 Res = GetSetting(gCompanyName & "Config", BookName, NameOrLabel(c), "NotFound")
9                 If Res <> "NotFound" Then
10                    If Not c.HasFormula Then
11                        SafeSetCellValue c, sParseArrayString(Res)(1, 1)
12                    End If
13                End If
14            Next c
15        Next a
16        Exit Sub
ErrHandler:
17        Throw "#GetConfigFromRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NameOrLabel
' Author    : Philip Swannell
' Date      : 08-Nov-2016
' Purpose   : If a cell has a name scoped to the sheet then returns the name, otherwise
'             returns the contents of the cell to the left. Plan to use this in
'             SaveConfigToRegistry and GetConfigFromRegistry
' -----------------------------------------------------------------------------------------------------------------------
Private Function NameOrLabel(c As Range)
          Dim Name As String
1         On Error GoTo ErrHandler
2         On Error Resume Next
3         Name = c.Name.Name
4         On Error GoTo ErrHandler
5         If InStr(Name, c.Parent.Name & "!") > 0 Or InStr(Name, "'" & c.Parent.Name & "'!") > 0 Then
6             If InStr(Name, c.Parent.Name) Then
7                 NameOrLabel = sStringBetweenStrings(Name, "!")
8                 Exit Function
9             End If
10        End If
11        NameOrLabel = CStr(c.Offset(0, -1).Value2)
12        Exit Function
ErrHandler:
13        Throw "#NameOrLabel (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
